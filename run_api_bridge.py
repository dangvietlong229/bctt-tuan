import os
import subprocess
import json
import pandas as pd
import sys

def main():
    workspace_dir = os.path.dirname(os.path.abspath(__file__))
    bridge_pull_js = os.path.join(workspace_dir, "api-fiin", "bridge_pull.js")
    output_json = os.path.join(workspace_dir, "api-fiin", "responses", "bridge_sample.json")

    print("=== 1. RUNNING NODE.JS SCRAPER BRIDGE ===")
    try:
        # Run node bridge_pull.js
        result = subprocess.run(
            ["node", bridge_pull_js],
            cwd=os.path.join(workspace_dir, "api-fiin"),
            capture_output=True,
            text=True,
            check=True
        )
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Error running node scraper: {e}")
        print("Stdout:", e.stdout)
        print("Stderr:", e.stderr)
        return 1
    except FileNotFoundError:
        print("Node.js is not installed or not in PATH. Please ensure Node.js is installed.")
        return 1

    print("\n=== 2. READING RETRIEVED DATA ===")
    if not os.path.exists(output_json):
        print(f"Error: Output file {output_json} was not created.")
        return 1

    try:
        with open(output_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"Error: Could not read valid bridge output: {exc}")
        return 1

    # 1. Display Market Liquidity (matchValue in Billion VND)
    print("\n[+] THANH KHOẢN VNINDEX (matchValue) - 5 NGÀY GẦN NHẤT:")
    vnindex_list = data.get("vnindexLiquidity", [])
    if vnindex_list:
        df_vnindex = pd.DataFrame(vnindex_list)
        if 'tradingDate' in df_vnindex.columns:
            df_vnindex['date'] = pd.to_datetime(df_vnindex['tradingDate'], format='ISO8601', utc=True).dt.date
        # Convert matchValue to Billion VND
        df_vnindex['matchValue_Bil'] = df_vnindex['matchValue'] / 1e9
        cols = ['date', 'closeIndex', 'matchValue_Bil', 'totalMatchVolume']
        cols = [c for c in cols if c in df_vnindex.columns]
        print(df_vnindex[cols].to_string(index=False))
    else:
        print("Không có dữ liệu thanh khoản.")

    # 2. Display Sector Index Alternative (VNREAL)
    print("\n[+] BIẾN ĐỘNG CHỈ SỐ NGÀNH BẤT ĐỘNG SẢN (VNREAL) - 5 NGÀY GẦN NHẤT:")
    vnreal_list = data.get("sectorVnReal", [])
    if vnreal_list:
        df_vnreal = pd.DataFrame(vnreal_list)
        if 'tradingDate' in df_vnreal.columns:
            df_vnreal['date'] = pd.to_datetime(df_vnreal['tradingDate'], format='ISO8601', utc=True).dt.date
        cols = ['date', 'closePrice', 'totalMatchVolume']
        cols = [c for c in cols if c in df_vnreal.columns]
        print(df_vnreal[cols].to_string(index=False))
    else:
        print("Không có dữ liệu ngành Bất động sản.")

    # 3. Display Foreign Trading for TCB (including YTD)
    print("\n[+] GIAO DỊCH KHỐI NGOẠI TCB (Foreign Stats including YTD) - ĐƠN VỊ TỶ ĐỒNG:")
    foreign_data = data.get("foreignInvestorTCB", {})
    if foreign_data:
        # Check if it has the standard structure
        items = foreign_data if isinstance(foreign_data, list) else [foreign_data]
        if items:
            item = items[0]
            summary_rows = []
            for period in ['today', 'oneWeek', 'oneMonth', 'yearToDate']:
                period_data = item.get(period, {})
                if period_data:
                    summary_rows.append({
                        'Period': period_data.get('timeRange'),
                        'From Date': pd.to_datetime(period_data.get('fromDate'), format='ISO8601', utc=True).date(),
                        'To Date': pd.to_datetime(period_data.get('toDate'), format='ISO8601', utc=True).date(),
                        'Buy Val (Bil)': round(period_data.get('foreignBuyValue', 0) / 1e9, 2),
                        'Sell Val (Bil)': round(period_data.get('foreignSellValue', 0) / 1e9, 2),
                        'Net Buy Val (Bil)': round(period_data.get('foreignNetBuyValue', 0) / 1e9, 2)
                    })
            df_summary = pd.DataFrame(summary_rows)
            print(df_summary.to_string(index=False))
    else:
        print("Không có dữ liệu khối ngoại.")

    # 4. Display Proprietary Trading VNINDEX (Market Summary)
    print("\n[+] GIAO DỊCH TỰ DOANH TOÀN THỊ TRƯỜNG (Proprietary Today Summary):")
    prop_market = data.get("proprietaryMarket", {})
    if prop_market:
        items = prop_market if isinstance(prop_market, list) else [prop_market]
        if items:
            item = items[0]
            today_data = item.get('today', {})
            if today_data:
                print(f"Tổng mua tự doanh hôm nay: {today_data.get('totalBuyTradeValue', 0) / 1e9:.2f} Tỷ VND")
                print(f"Tổng bán tự doanh hôm nay: {today_data.get('totalSellTradeValue', 0) / 1e9:.2f} Tỷ VND")
                # Show top 3 stocks bought by proprietary trading today
                buy_stocks = today_data.get('buy', [])
                if buy_stocks:
                    print("\nTop 3 cổ phiếu được Tự doanh mua nhiều nhất hôm nay:")
                    df_buy = pd.DataFrame(buy_stocks).sort_values(by='totalBuyTradeValue', ascending=False).head(3)
                    df_buy['Buy Val (Bil)'] = df_buy['totalBuyTradeValue'] / 1e9
                    cols = ['ticker', 'Buy Val (Bil)', 'percentPriceChange']
                    cols = [c for c in cols if c in df_buy.columns]
                    print(df_buy[cols].to_string(index=False))
    else:
        print("Không có dữ liệu tự doanh toàn thị trường.")

    print("\n=== HOÀN THÀNH KIỂM TRA DỮ LIỆU ===")
    return 0

if __name__ == '__main__':
    sys.exit(main())
