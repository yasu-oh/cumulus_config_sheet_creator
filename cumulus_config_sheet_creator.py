import yaml
import pandas as pd
from collections import defaultdict
import sys
import os

def load_yaml_file(file_path):
    """YAMLファイルを読み込む"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        print(f"YAMLファイルの読み込みエラー: {e}")
        sys.exit(1)

def extract_snmp_trap_destinations(data):
    """SNMP trap destinationを抽出"""
    trap_destinations = []

    if 'system' in data[1]['set'] and 'snmp-server' in data[1]['set']['system']:
        snmp_config = data[1]['set']['system']['snmp-server']

        # trap-destinationの処理
        if 'trap-destination' in snmp_config:
            trap_destinations_dict = snmp_config['trap-destination']

            for dest_ip, dest_config in trap_destinations_dict.items():
                if 'vrf' in dest_config and 'mgmt' in dest_config['vrf']:
                    vrf_config = dest_config['vrf']['mgmt']
                    if 'community-password' in vrf_config:
                        community_passwords = vrf_config['community-password']
                        for community_name, community_config in community_passwords.items():
                            trap_destinations.append({
                                'destination': dest_ip,
                                'community_password': community_name
                            })

    return trap_destinations

def extract_interface_settings(data):
    """インタフェース設定を抽出（bondインタフェース対応）"""
    interfaces = defaultdict(dict)

    # bondインタフェース設定を抽出
    if 'interface' in data[1]['set']:
        for key, value in data[1]['set']['interface'].items():
            # bondインタフェースの情報を抽出
            if 'bond' in value:
                bond_config = value['bond']
                interfaces[key] = {
                    'type': 'bond',
                    'description': value.get('description', ''),
                    'ip_address': list(value.get('ip', {}).get('address', {}).keys()) if value.get('ip') else [],
                    'vrf': value.get('ip', {}).get('vrf', ''),
                    'bond_mode': bond_config.get('mode', ''),
                    'lacp_rate': bond_config.get('lacp-rate', ''),
                    'lacp_bypass': bond_config.get('lacp-bypass', ''),
                    'members': list(bond_config.get('member', {}).keys()) if bond_config.get('member') else [],
                    'mtu': value.get('link', {}).get('mtu', ''),
                    'bridge_domain': list(value.get('bridge', {}).get('domain', {}).keys()) if value.get('bridge', {}).get('domain') else [],
                    'access_vlan': list(list(value.get('bridge', {}).get('domain', {}).values())[0].values())[0] if value.get('bridge', {}).get('domain') and list(value.get('bridge', {}).get('domain', {}).values()) and list(value.get('bridge', {}).get('domain', {}).values())[0].get('access') else '',
                    'stp_admin_edge': list(list(value.get('bridge', {}).get('domain', {}).values())[0].get('stp', {}).values())[0] if value.get('bridge', {}).get('domain') and list(value.get('bridge', {}).get('domain', {}).values()) and list(value.get('bridge', {}).get('domain', {}).values())[0].get('stp') else '',
                    'stp_auto_edge': list(list(value.get('bridge', {}).get('domain', {}).values())[0].get('stp', {}).values())[1] if value.get('bridge', {}).get('domain') and list(value.get('bridge', {}).get('domain', {}).values()) and len(list(value.get('bridge', {}).get('domain', {}).values())[0].get('stp', {})) >= 2 else '',
                    'stp_bpdu_guard': list(list(value.get('bridge', {}).get('domain', {}).values())[0].get('stp', {}).values())[2] if value.get('bridge', {}).get('domain') and list(value.get('bridge', {}).get('domain', {}).values()) and len(list(value.get('bridge', {}).get('domain', {}).values())[0].get('stp', {})) >= 3 else '',
                }
            else:
                # 物理インタフェースの情報を抽出
                interfaces[key] = {
                    'type': value.get('type', ''),
                    'description': value.get('description', ''),
                    'ip_address': list(value.get('ip', {}).get('address', {}).keys()) if value.get('ip') else [],
                    'vrf': value.get('ip', {}).get('vrf', ''),
                    'breakout': '',
                    'mtu': value.get('link', {}).get('mtu', ''),
                }
                # Breakout設定の抽出
                if 'link' in value and 'breakout' in value['link']:
                    interfaces[key]['breakout'] = list(value['link']['breakout'])[0]

    return interfaces

def extract_vrf_bgp_settings(data):
    """VRFごとのBGP設定を抽出"""
    vrf_bgp_configs = {}

    # VRFのルーターパスをたどる
    if 'vrf' in data[1]['set']:
        for vrf_name, vrf_data in data[1]['set']['vrf'].items():
            if 'router' in vrf_data and 'bgp' in vrf_data['router']:
                bgp_config = vrf_data['router']['bgp']

                # ネイバー情報を抽出
                neighbors = {}
                if 'neighbor' in bgp_config:
                    for ip, neighbor_info in bgp_config['neighbor'].items():
                        neighbors[ip] = {
                            'peer-group': neighbor_info.get('peer-group', ''),
                            'type': neighbor_info.get('type', ''),
                            'remote-as': neighbor_info.get('remote-as', ''),
                        }

                # ピアグループ情報を抽出
                peer_groups = {}
                if 'peer-group' in bgp_config:
                    for pg_name, pg_info in bgp_config['peer-group'].items():
                        peer_groups[pg_name] = {
                            'description': pg_info.get('description', ''),
                            'remote-as': pg_info.get('remote-as', ''),
                            'update-source': pg_info.get('update-source', ''),
                            'bfd-enable': pg_info.get('bfd', {}).get('enable', False),
                            'bfd-min-rx-interval': pg_info.get('bfd', {}).get('min-rx-interval', ''),
                            'bfd-min-tx-interval': pg_info.get('bfd', {}).get('min-tx-interval', ''),
                            'bfd-detect-multiplier': pg_info.get('bfd', {}).get('detect-multiplier', ''),
                            'multihop-ttl': pg_info.get('multihop-ttl', ''),
                        }

                # アドレスファミリ情報を抽出
                address_families = {}
                if 'address-family' in bgp_config:
                    for af_name, af_info in bgp_config['address-family'].items():
                        address_families[af_name] = {
                            'enable': af_info.get('enable', False),
                            'redistribute': af_info.get('redistribute', {}),
                            'route-export': af_info.get('route-export', {}),
                        }

                # パス選択情報を抽出
                path_selection = {}
                if 'path-selection' in bgp_config:
                    path_selection = bgp_config['path-selection']

                vrf_bgp_configs[vrf_name] = {
                    'router-id': bgp_config.get('router-id', ''),
                    'autonomous-system': bgp_config.get('autonomous-system', ''),
                    'enable': bgp_config.get('enable', False),
                    'neighbors': neighbors,
                    'peer-groups': peer_groups,
                    'address-families': address_families,
                    'path-selection': path_selection,
                }

    return vrf_bgp_configs

def create_interface_dataframe(interfaces):
    """インタフェース情報をDataFrameに変換"""
    return pd.DataFrame([
        {
            'Interface': name,
            'Type': info.get('type', ''),
            'Breakout': info.get('breakout', ''),
            'Description': info.get('description', ''),
            'IP Address': ', '.join(info.get('ip_address', [])),
            'VRF': info.get('vrf', ''),
            'Bond Mode': info.get('bond_mode', ''),
            'LACP Rate': info.get('lacp_rate', ''),
            'LACP Bypass': info.get('lacp_bypass', ''),
            'Members': ', '.join(info.get('members', [])),
            'MTU': info.get('mtu', ''),
            'Bridge Domain': ', '.join(info.get('bridge_domain', [])),
            'Access VLAN': info.get('access_vlan', ''),
            'STP Admin Edge': info.get('stp_admin_edge', ''),
            'STP Auto Edge': info.get('stp_auto_edge', ''),
            'STP BPDU Guard': info.get('stp_bpdu_guard', '')
        }
        for name, info in interfaces.items()
    ])

def create_bgp_vrf_dataframe(vrf_bgp_configs):
    """VRFごとのBGP設定をDataFrameに変換"""
    rows = []

    for vrf_name, bgp_config in vrf_bgp_configs.items():
        # パス選択情報の抽出
        multipath_info = ""
        if 'multipath' in bgp_config.get('path-selection', {}):
            multipath_config = bgp_config['path-selection']['multipath']
            multipath_info = f"aspath-ignore: {multipath_config.get('aspath-ignore', 'off')}"

        rows.append({
            'VRF Name': vrf_name,
            'Router ID': bgp_config.get('router-id', ''),
            'Autonomous System': bgp_config.get('autonomous-system', ''),
            'Enable': bgp_config.get('enable', False),
            'Neighbors Count': len(bgp_config.get('neighbors', {})),
            'Peer Groups Count': len(bgp_config.get('peer-groups', {})),
            'Multipath Info': multipath_info,
        })

    return pd.DataFrame(rows)

def create_bgp_neighbors_dataframe(vrf_bgp_configs):
    """VRFごとのBGPネイバー情報をDataFrameに変換"""
    rows = []

    for vrf_name, bgp_config in vrf_bgp_configs.items():
        neighbors = bgp_config.get('neighbors', {})
        for ip, neighbor_info in neighbors.items():
            rows.append({
                'VRF Name': vrf_name,
                'Neighbor IP': ip,
                'Peer Group': neighbor_info.get('peer-group', ''),
                'Type': neighbor_info.get('type', ''),
                'Remote AS': neighbor_info.get('remote-as', ''),
            })

    return pd.DataFrame(rows)

def create_bgp_peer_groups_dataframe(vrf_bgp_configs):
    """VRFごとのBGPピアグループ情報をDataFrameに変換"""
    rows = []

    for vrf_name, bgp_config in vrf_bgp_configs.items():
        peer_groups = bgp_config.get('peer-groups', {})
        for pg_name, pg_info in peer_groups.items():
            rows.append({
                'VRF Name': vrf_name,
                'Peer Group': pg_name,
                'Description': pg_info.get('description', ''),
                'Remote AS': pg_info.get('remote-as', ''),
                'Update Source': pg_info.get('update-source', ''),
                'BFD Enable': pg_info.get('bfd-enable', False),
                'BFD Min Rx Interval': pg_info.get('bfd-min-rx-interval', ''),
                'BFD Min Tx Interval': pg_info.get('bfd-min-tx-interval', ''),
                'BFD Detect Multiplier': pg_info.get('bfd-detect-multiplier', ''),
                'Multihop TTL': pg_info.get('multihop-ttl', ''),
            })

    return pd.DataFrame(rows)

def create_other_settings_dataframe(data, trap_destinations):
    """その他の設定情報をDataFrameに変換"""
    # モデルとバージョンを抽出
    model = ""
    version = ""
    if 'header' in data[0]:
        model = data[0]['header'].get('model', '')
        version = data[0]['header'].get('version', '')

    # SNMPコミュニティを抽出
    snmp_communities = []
    snmp_community_access = {}

    if 'set' in data[1] and 'system' in data[1]['set'] and 'snmp-server' in data[1]['set']['system']:
        snmp_config = data[1]['set']['system']['snmp-server']

        # readonly-communityを抽出
        if 'readonly-community' in snmp_config:
            for community_name, community_info in snmp_config['readonly-community'].items():
                snmp_communities.append(community_name)

                # アクセス制御情報を抽出
                if 'access' in community_info:
                    access_ips = list(community_info['access'].keys())
                    snmp_community_access[community_name] = access_ips

    # NTPサーバーを抽出
    ntp_servers = []
    if 'set' in data[1] and 'service' in data[1]['set'] and 'ntp' in data[1]['set']['service']:
        ntp_config = data[1]['set']['service']['ntp']
        if 'mgmt' in ntp_config:
            ntp_servers = list(ntp_config['mgmt'].get('server', {}).keys())

    # DNSサーバーを抽出
    dns_servers = []
    if 'set' in data[1] and 'service' in data[1]['set'] and 'dns' in data[1]['set']['service']:
        dns_config = data[1]['set']['service']['dns']
        if 'mgmt' in dns_config:
            dns_servers = list(dns_config['mgmt'].get('server', {}).keys())

    # syslogサーバーを抽出
    syslog_servers = []
    if 'set' in data[1] and 'service' in data[1]['set'] and 'syslog' in data[1]['set']['service']:
        syslog_config = data[1]['set']['service']['syslog']
        if 'mgmt' in syslog_config:
            syslog_servers = list(syslog_config['mgmt'].get('server', {}).keys())

    # ブリッジ設定を抽出
    bridge_stp_priority = ""
    bridge_vlans = []
    if 'set' in data[1] and 'bridge' in data[1]['set']:
        bridge_config = data[1]['set']['bridge']
        if 'domain' in bridge_config:
            domain_config = bridge_config['domain']
            for domain_name, domain_info in domain_config.items():
                if 'stp' in domain_info:
                    stp_config = domain_info['stp']
                    bridge_stp_priority = stp_config.get('priority', '')

                if 'vlan' in domain_info:
                    vlan_config = domain_info['vlan']
                    bridge_vlans = list(vlan_config.keys())

    # EVPN設定を抽出
    evpn_enable = ""
    multihoming_enable = ""
    if 'set' in data[1] and 'evpn' in data[1]['set']:
        evpn_config = data[1]['set']['evpn']
        evpn_enable = evpn_config.get('enable', '')
        if 'multihoming' in evpn_config:
            multihoming_enable = evpn_config['multihoming'].get('enable', '')

    # ホスト名を抽出
    hostname = ""
    if 'set' in data[1] and 'system' in data[1]['set'] and 'hostname' in data[1]['set']['system']:
        hostname = data[1]['set']['system']['hostname']

    other_settings = {
        'Hostname': hostname,
        'Model': model,
        'Version': version,
        'Router ID': data[1]['set'].get('router', {}).get('bgp', {}).get('router-id'),
        'Autonomous System': data[1]['set'].get('router', {}).get('bgp', {}).get('autonomous-system'),
        'SNMP Community': ', '.join(snmp_communities),
        'SNMP Trap Destinations': ', '.join([f"{dest['destination']} -> {dest['community_password']}" for dest in trap_destinations]),
        'NTP Server': ', '.join(ntp_servers),
        'DNS Server': ', '.join(dns_servers),
        'Syslog Servers': ', '.join(syslog_servers),
        'Bridge STP Priority': bridge_stp_priority,
        'Bridge VLANs': ', '.join(bridge_vlans),
        'EVPN Enable': evpn_enable,
        'Multihoming Enable': multihoming_enable,
    }

    return pd.DataFrame([other_settings])

def main():
    # YAMLファイルのパスを指定
    yaml_file_path = sys.argv[1]

    # ファイル存在確認
    if not os.path.exists(yaml_file_path):
        print(f"エラー: ファイルが見つかりません - {yaml_file_path}")
        sys.exit(1)

    # YAMLファイルを読み込む
    data = load_yaml_file(yaml_file_path)

    # インタフェース設定の抽出
    interfaces = extract_interface_settings(data)

    # VRFごとのBGP設定の抽出
    vrf_bgp_configs = extract_vrf_bgp_settings(data)

    # SNMP trap destinationの抽出
    trap_destinations = extract_snmp_trap_destinations(data)

    # DataFrameに変換
    interface_df = create_interface_dataframe(interfaces)
    bgp_vrf_df = create_bgp_vrf_dataframe(vrf_bgp_configs)
    bgp_neighbors_df = create_bgp_neighbors_dataframe(vrf_bgp_configs)
    bgp_peer_groups_df = create_bgp_peer_groups_dataframe(vrf_bgp_configs)
    other_settings_df = create_other_settings_dataframe(data, trap_destinations)
    other_settings_kv_df = other_settings_df.melt(var_name='Key', value_name='Value')

    # ホスト名を取得
    hostname = other_settings_df.iloc[0]['Hostname']

    # 出力ファイル名を決定
    if hostname:
        output_file = f"{hostname}.xlsx"
    else:
        output_file = "cumulus_configuration_sheet.xlsx"

    # エクセルファイルに出力
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        other_settings_kv_df.to_excel(writer, sheet_name='General Settings', index=False)
        interface_df.to_excel(writer, sheet_name='Interface Settings', index=False)
        bgp_vrf_df.to_excel(writer, sheet_name='BGP VRF Summary', index=False)
        bgp_neighbors_df.to_excel(writer, sheet_name='BGP Neighbors', index=False)
        bgp_peer_groups_df.to_excel(writer, sheet_name='BGP Peer Groups', index=False)

    print(f"エクセルファイルを出力しました: {output_file}")

if __name__ == '__main__':
    main()
