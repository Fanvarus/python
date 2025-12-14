import requests
import pandas as pd
import numpy as np
from datetime import datetime

# 配置账号信息
ACCOUNTS = {
    "jiweishidai": "6D218509562ED94DB2808E28AE3DB3BB",
    "huangfangyi": "6F0A6BC78A79D8E922410BB0971FDE0A"
}

# 接口基础配置
BASE_URL = "https://sapi.musmoon.com"
LOGIN_URL = f"{BASE_URL}/card/user/password/login"
BALANCE_URL = f"{BASE_URL}/card/proxy/company/capital/account/info?currencyType=CNY"
BILL_URL = f"{BASE_URL}/card/proxy/user/bill/page?currency=CNY&billType=&orderNo=&cardValue=&orders[0].column=id&orders[0].asc=false&current=1&size=10"

# 账单字段中文翻译映射（覆盖全部字段）
BILL_FIELD_TRANSLATE = {
    "orderNo": "订单号",
    "capitalAccountType": "资金账户类型",
    "billType": "账单类型",
    "uid": "用户ID",
    "billAmount": "交易金额",
    "beforeAmount": "交易前余额",
    "cardIccid": "卡ICCID码",
    "createTime": "创建时间",
    "currency": "币种",
    "id": "账单ID",
    "afterAmount": "交易后余额",
    "cardNumber": "卡号",
    "remarks": "备注",
    # 扩展兼容其他可能的账单字段（备用）
    "updateTime": "更新时间",
    "companyId": "企业ID",
    "companyUid": "企业用户ID",
    "riskAmount": "风险冻结金额",
    "creditAmount": "授信额度"
}

# 账单类型中文翻译
BILL_TYPE_TRANSLATE = {
    "orderCommissionBill": "订单佣金账单",
    "orderRefundBill": "订单退款账单",
    "userWithdraw": "用户提现账单",
    "unknown": "未知账单类型"
}

# 全局变量（保留所有原始数据备用）
all_balance_raw_data = {}  # 余额原始全字段 {账号: 全字段字典}
all_bill_raw_data = {}  # 账单原始全字段 {账号: [账单1全字段, 账单2全字段...]}
total_balance = 0.0  # 总余额（可提现+不可提现）
total_withdrawable = 0.0  # 总可提现余额
total_non_withdrawable = 0.0  # 总不可提现余额


def login(username: str, password: str) -> str | None:
    """登录接口，获取token"""
    try:
        login_params = {"username": username, "password": password}
        response = requests.post(LOGIN_URL, params=login_params, timeout=10)
        response.raise_for_status()

        result = response.json()
        if result.get("success") and result.get("statusCode") == 0:
            token = result["object"]["token"]
            print(f"\n=====================================")
            print(f"✅ 账号【{username}】登录成功")
            print(f"🔑 Token：{token}")
            print(f"=====================================")
            return token
        else:
            print(f"\n❌ 账号【{username}】登录失败：{result.get('content', '未知错误')}")
            return None
    except Exception as e:
        print(f"\n❌ 账号【{username}】登录异常：{str(e)}")
        return None


def get_balance(token: str, username: str) -> None:
    """
    余额接口：读取全部字段，重构显示逻辑
    显示规则：
    1. 总余额 = 可提现余额 + 不可提现余额
    2. 显示可提现余额
    3. 若不可提现余额>0，显示不可提现金额
    4. 所有字段保留到全局变量备用
    """
    global total_balance, total_withdrawable, total_non_withdrawable
    try:
        headers = {"x-token": f'{{"token":"{token}"}}'}
        response = requests.get(BALANCE_URL, headers=headers, timeout=10)
        response.raise_for_status()

        # 读取余额全部字段并保存备用
        balance_raw = response.json()
        all_balance_raw_data[username] = balance_raw
        print(f"\n📜 账号【{username}】余额接口全字段原始数据：")
        print(f"   {balance_raw}")

        if balance_raw.get("success") and balance_raw.get("statusCode") == 0:
            balance_info = balance_raw.get("object", {})

            # 核心字段解析
            withdrawable = float(balance_info.get("withdrawAmount", 0.0))  # 可提现余额
            non_withdrawable = float(balance_info.get("nonWithdrawAmount", 0.0))  # 不可提现余额
            # 兜底：若没有nonWithdrawAmount，用balance - withdrawable计算
            if non_withdrawable == 0.0 and "balance" in balance_info:
                non_withdrawable = float(balance_info["balance"]) - withdrawable

            total = withdrawable + non_withdrawable  # 总余额

            # 累加至全局汇总
            total_balance += total
            total_withdrawable += withdrawable
            total_non_withdrawable += non_withdrawable

            # 格式化显示
            print(f"\n💰 账号【{username}】余额核心信息")
            print(f"   总余额：{total:.2f} 元（可提现 {withdrawable:.2f} 元）")
            if non_withdrawable > 0:
                print(f"   不可提现余额：{non_withdrawable:.2f} 元")
        else:
            print(f"\n❌ 账号【{username}】余额查询失败：{balance_raw.get('content', '未知错误')}")
    except Exception as e:
        print(f"\n❌ 账号【{username}】余额查询异常：{str(e)}")


def translate_bill_record(raw_record: dict) -> dict:
    """翻译单条账单的全部字段"""
    translated = {}
    for en_key, value in raw_record.items():
        # 翻译字段名（无映射则保留原字段名）
        cn_key = BILL_FIELD_TRANSLATE.get(en_key, en_key)

        # 翻译字段值
        if en_key == "billType":
            translated[cn_key] = BILL_TYPE_TRANSLATE.get(value, value)
        elif en_key == "currency" and value == "CNY":
            translated[cn_key] = "人民币"
        elif en_key in ["billAmount", "beforeAmount", "afterAmount"] and value is not None:
            translated[cn_key] = f"{float(value):.2f} 元"
        elif value is None:
            translated[cn_key] = "无"
        else:
            translated[cn_key] = value
    return translated


def get_and_print_bill(token: str, username: str) -> None:
    """
    账单接口：读取全部字段，打印核心明细，保留所有字段备用
    """
    global all_bill_raw_data
    try:
        headers = {"x-token": f'{{"token":"{token}"}}'}
        response = requests.get(BILL_URL, headers=headers, timeout=10)
        response.raise_for_status()

        # 读取账单全部字段并保存备用
        bill_raw = response.json()
        all_bill_raw_data[username] = bill_raw
        bill_records = bill_raw.get("object", {}).get("records", [])

        print(f"\n📜 账号【{username}】账单接口全字段原始数据（总条数：{bill_raw.get('object', {}).get('total', 0)}）：")
        print(f"   接口返回全字段：{bill_raw}")

        if not bill_records:
            print(f"\n📃 账号【{username}】无账单数据")
            return

        # 逐条打印账单原始全字段+翻译
        print(f"\n📝 账号【{username}】账单逐条解析（共{len(bill_records)}条）：")
        for idx, raw_rec in enumerate(bill_records, 1):
            print(f"\n   第{idx}条账单原始全字段：")
            print(f"      {raw_rec}")
            translated_rec = translate_bill_record(raw_rec)
            print(f"   第{idx}条账单翻译后：")
            print(f"      {translated_rec}")
            print("   " + "-" * 100)

        # 核心明细表格（保留原有格式）
        df = pd.DataFrame(bill_records)
        # 确保核心字段存在
        for field, default in {"id": "未知ID", "billAmount": 0.0, "createTime": "未知时间",
                               "orderNo": "无订单号"}.items():
            if field not in df.columns:
                df[field] = default

        df["orderNo"] = df["orderNo"].fillna("无订单号")
        df["createTime"] = df["createTime"].fillna("未知时间")
        df["billAmount"] = pd.to_numeric(df["billAmount"], errors="coerce").fillna(0.0)
        df["收支类型"] = df["billAmount"].apply(lambda x: "收入" if x > 0 else ("支出" if x < 0 else "无变动"))
        df["所属账号"] = username

        # 核心表格展示
        core_df = df[["所属账号", "id", "billAmount", "收支类型", "createTime", "orderNo"]].rename(columns={
            "id": "账单ID",
            "billAmount": "交易金额(元)",
            "createTime": "交易时间",
            "orderNo": "订单号"
        }).reset_index(drop=True)
        core_df = core_df[core_df["交易金额(元)"].abs() <= 10000]

        print(f"\n📋 账号【{username}】账单核心明细表格：")
        print("-" * 120)
        print(f"{'所属账号':<12}{'账单ID':<12}{'交易金额(元)':<15}{'收支类型':<8}{'交易时间':<22}{'订单号'}")
        print("-" * 120)
        for _, row in core_df.iterrows():
            print(
                f"{str(row['所属账号']):<12}{str(row['账单ID']):<12}{float(row['交易金额(元)']):<15.2f}{str(row['收支类型']):<8}{str(row['交易时间']):<22}{str(row['订单号'])}")

        # 单账号账单汇总
        income = core_df[core_df["收支类型"] == "收入"]["交易金额(元)"].sum()
        expense = core_df[core_df["收支类型"] == "支出"]["交易金额(元)"].sum()
        print("-" * 120)
        print(
            f"📊 账号【{username}】账单汇总：收入 {income:.2f} 元 | 支出 {expense:.2f} 元 | 净收支 {income + expense:.2f} 元")

    except Exception as e:
        print(f"\n❌ 账号【{username}】账单查询异常：{str(e)}")


def print_total_summary():
    """全局汇总信息"""
    print(f"\n=====================================")
    print(f"📈 所有账号汇总信息")
    print(f"=====================================")
    print(f"💰 余额汇总：")
    print(f"   总余额：{total_balance:.2f} 元")
    print(f"   总可提现余额：{total_withdrawable:.2f} 元")
    if total_non_withdrawable > 0:
        print(f"   总不可提现余额：{total_non_withdrawable:.2f} 元")

    # 账单汇总
    all_bills = []
    for username, bill_data in all_bill_raw_data.items():
        records = bill_data.get("object", {}).get("records", [])
        df = pd.DataFrame(records)
        if not df.empty:
            df["billAmount"] = pd.to_numeric(df["billAmount"], errors="coerce").fillna(0.0)
            all_bills.append(df)

    if all_bills:
        total_df = pd.concat(all_bills, ignore_index=True)
        total_income = total_df[total_df["billAmount"] > 0]["billAmount"].sum()
        total_expense = total_df[total_df["billAmount"] < 0]["billAmount"].sum()
        print(f"\n📃 账单汇总：")
        print(f"   总账单条数：{len(total_df)} 条")
        print(f"   总收入：{total_income:.2f} 元")
        print(f"   总支出：{total_expense:.2f} 元")
        print(f"   总净收支：{total_income + total_expense:.2f} 元")
    else:
        print(f"\n📃 账单汇总：无账单数据")
    print(f"=====================================")


def main():
    """主流程"""
    for username, password in ACCOUNTS.items():
        token = login(username, password)
        if not token:
            continue

        # 余额查询（全字段+重构显示）
        get_balance(token, username)

        # 账单查询（全字段+翻译+核心表格）
        get_and_print_bill(token, username)

    # 全局汇总
    print_total_summary()

    # 可选：打印备用数据的存储提示
    print(f"\n💾 备用数据说明：")
    print(f"   - 余额全字段已保存至 all_balance_raw_data 字典（key=账号名）")
    print(f"   - 账单全字段已保存至 all_bill_raw_data 字典（key=账号名）")


if __name__ == "__main__":
    # pip install requests pandas numpy
    main()