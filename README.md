# Cumulus Linux 設計シート生成ツール

Cumulus Linux / NVUE の YAML 形式構成を解析し、インタフェース・BGP・EVPN・SNMP などの主要項目を **Excel（.xlsx）設計シート** に自動整形して出力します。

- 入力: NVUE の `nv config show` などで取得した **YAML**
- 出力: ホスト名（取得できた場合）に基づく `HOSTNAME.xlsx` もしくは `cumulus_configuration_sheet.xlsx`
- 対象: Bond（LAG/LACP）・物理 IF・BGP（VRF/AFI/ネイバー/ピアグループ）・SNMP・NTP/DNS/Syslog・Bridge（STP/VLAN）・EVPN（マルチホーミング）他

## 機能
- NVUE YAML から以下の設定を抽出して表形式に整形
  - **Interface Settings**: 物理 / Bond、説明、IP、VRF、MTU、Bridge/Access VLAN、STP（Admin/Auto/BPDU Guard）
  - **BGP VRF Summary**: VRF 単位の AS、Router-ID、Multipath ポリシ、ネイバー/PG 数
  - **BGP Neighbors**: VRF、Neighbor IP、Peer Group、Type、Remote AS
  - **BGP Peer Groups**: VRF、説明、Remote-AS、update-source、BFD（enable/interval/multiplier）、Multihop TTL
  - **Other Settings**: Hostname、モデル/バージョン、グローバル BGP（Router-ID/AS）、SNMP readonly-community とアクセス許可 IP、Trap 宛先、NTP/DNS/Syslog、Bridge STP 優先度/VLAN、EVPN/Multihoming 有効化
- **ホスト名を自動検出** し、出力ファイル名に反映
- Excel へ **複数シート** で出力（`openpyxl`）

## 前提条件
- Python 3.9+ 推奨（3.8 以上で動作想定）
- 依存パッケージ
  - `PyYAML`
  - `pandas`
  - `openpyxl`

## インストール
```bash
pip install pyyaml pandas openpyxl
```

## 使い方
```bash
python cumulus_config_sheet_creator.py /path/to/running-config.yaml
```
- 正常終了時、`<Hostname>.xlsx`（Hostname が未取得の場合は `cumulus_configuration_sheet.xlsx`）をカレントディレクトリに生成します。

## 出力ファイルとシート構成
- **Interface Settings**
  - 物理/Bond の各 IF を 1 行で記述。Bond はメンバ一覧や LACP 設定、Bridge/Access VLAN、STP 状態を含みます。
- **BGP VRF Summary**
  - VRF ごとの BGP 要約（AS、Router-ID、Enable、Multipath 設定要約、Neighbor/PeerGroup 数）。
- **BGP Neighbors**
  - 全 VRF のネイバー一覧と属性。
- **BGP Peer Groups**
  - 全 VRF のピアグループ一覧と詳細（BFD/TTL/US 等）。
- **Other Settings**
  - Hostname、Model、Version、グローバル BGP（`set.router.bgp`）の Router-ID/AS、SNMP（readonly-community とアクセス IP、Trap 宛先）、NTP/DNS/Syslog、Bridge STP Priority/VLAN、EVPN 有効化など。

## データ項目の定義
### Interface Settings
| 列 | 説明 |
|---|---|
| Interface | IF 名（例: `swp1`, `bond1`） |
| Type | `bond` または物理種別（空の場合あり） |
| Breakout | `link.breakout` の設定（存在時） |
| Description | `description` |
| IP Address | `ip.address` の複数値をカンマ区切りで列挙 |
| VRF | `ip.vrf` |
| Bond Mode | `bond.mode`（例: `802.3ad`） |
| LACP Rate | `bond.lacp-rate`（`fast` 等） |
| LACP Bypass | `bond.lacp-bypass` |
| Members | `bond.member` のキー一覧 |
| MTU | `link.mtu` |
| Bridge Domain | `bridge.domain` 名 |
| Access VLAN | `bridge.domain.<domain>.access` |
| STP Admin Edge / Auto Edge / BPDU Guard | `bridge.domain.<domain>.stp` 各項目 |

### BGP VRF Summary / Neighbors / Peer Groups
- NVUE `vrf.<name>.router.bgp` 配下を解析し、VRF 単位で要約化/正規化して出力。
- Multipath の要約は `path-selection.multipath.aspath-ignore` を `aspath-ignore: on|off` 形式で記載。

### Other Settings
- `header.model` / `header.version`
- `set.system.hostname`
- `set.router.bgp.router-id` / `autonomous-system`
- `set.system.snmp-server.readonly-community.*.access`（許可 IP を列挙）
- `set.system.snmp-server.trap-destination.*.vrf.mgmt.community-password.*`（宛先→コミュニティ名）
- `set.service.ntp.mgmt.server.*` / `set.service.dns.mgmt.server.*` / `set.service.syslog.mgmt.server.*`
- `set.bridge.domain.*.stp.priority` / `set.bridge.domain.*.vlan.*`
- `set.evpn.enable` / `set.evpn.multihoming.enable`

## 制限事項・既知の注意点
- **YAML スキーマ**: 本スクリプトは NVUE の一般的な YAML 構造を前提とします。旧 ifupdown2 構成、あるいはベンダー拡張により階層が異なる場合は抽出関数の微修正が必要です。
- **Bridge/STP の抽出**: `bridge.domain` が複数ある場合、各 IF に最初のドメインの属性を割当てています。必要に応じて *IF×Domain* の正規化に拡張してください。
- **SNMP Trap**: `community-password` キーは実体としてコミュニティ名を採っています（命名が紛らわしい点に留意）。秘匿情報の扱いには十分注意してください。
- **アドレスファミリ**: `address-family` の詳細（例えば `l2vpn-evpn` の redistribute/export ルール等）は現状要約のみ。列挙を追加する場合は `create_bgp_vrf_dataframe()` の拡張が必要です。

## 拡張のヒント
- **VLAN/VRF/IF の整合チェック**: VLAN 存在確認、IF⇔VRF の紐付け検証、未使用 VLAN の抽出などの QA ルールを追加。
- **EVPN MH 詳細**: `evpn.multihoming.*`（ESI/LACP mode/MH timer）や `nv show evpn multihoming` の出力を別シート化。
- **BGP AFI/SAFI 詳細**: `address-family.l2vpn-evpn`／`ipv4-unicast` の `redistribute` 等を表展開。
- **ACL/CoPP/Route-Map**: `policy.*` を解析して別シートに展開。
- **単体テスト**: サンプル YAML を fixtures 化し、`pytest` で関数別に検証。
