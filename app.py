import pandas as pd
import numpy as np


credit_memo_number = ""


class ReadCreditMemoRaport:
    def read_credit_memo_raport(self):
        credit_memo_raport = pd.read_excel("cm.xlsx")
        credit_memo_number = credit_memo_raport["POS Cr Memo #"]
        credit_memo_number.drop_duplicates(
            keep="first", inplace=True, ignore_index=True
        )
        # print(credit_memo_number)
        return credit_memo_raport


class ReadSalesRaport:
    def read_sales_raport(self):
        sales_raport = pd.read_excel("cmo.xls", skiprows=4)
        return sales_raport


class JoinRaports:

    def __init__(self, credit_memo, sales_raport):
        self.credit_memo = credit_memo
        self.sales_raport = sales_raport

    def join_raports(self):
        show_columns = [
            "Partner name",
            "Ship to Country",
            "Customer name",
            "Customer Code",
            "Invoice Number",
            "Spi Number",
            "Transaction Data",
            "SanDisk Part Number",
            "Delivery document number",
            "Foreign document",
            "Exchange rate EUR",
            "Exchange rate USD",
            "Customs number",
            "Sales Quantity",
            "Selling price",
            "Purchase price",
            "Unit Rebate (DC)",
            "Invoice",
        ]
        rename_columns = {
            "Ship to Country": "country ",
            "Delivery document number": "Numer dokumentu",
            "Foreign document": "Dokument obcy",
            "Exchange rate EUR": "Kurs EUR",
            "Exchange rate USD": "Kurs USD",
            "Customs number": "Kod celny",
            "Selling price": "Selling price: jedn. cena sprzedaży",
            "Purchase price": "Purchase price: jedn. cena zakupu",
            "Unit Rebate (DC)": "rabat",
        }

        # ========================================================================

        #

        credit_memo_number = credit_memo.loc[:0, ["POS Cr Memo #"]]
        credit_memo_number = credit_memo_number.to_string(header=False, index=False)

        merged_reports_1 = pd.merge(
            sales_raport,
            credit_memo,
            left_on=["Invoice Number", "SanDisk Part Number"],
            right_on=["Invoice", "Debit MPN"],
        )

        merged_reports_2 = pd.merge(
            sales_raport,
            credit_memo,
            left_on=["Spi Number", "SanDisk Part Number"],
            right_on=["Invoice", "Debit MPN"],
        )

        merged_reports_3 = merged_reports_2._append(merged_reports_1, ignore_index=True)

        # merged_reports_1.to_excel("Merged_1.xlsx")
        # merged_reports_2.to_excel("Merged_2.xlsx")

        merged_reports_3 = merged_reports_3[show_columns]

        merged_reports_3.rename(columns=rename_columns, inplace=True)

        merged_reports_3["USD"] = (
            merged_reports_3["Sales Quantity"] * merged_reports_3["rabat"]
        )
        merged_reports_3["Kurs USD"] = [
            float(str(i).replace(",", ".")) for i in merged_reports_3["Kurs USD"]
        ]
        merged_reports_3["PLN"] = merged_reports_3["USD"] * merged_reports_3["Kurs USD"]
        merged_reports_3["wartość sprzedaży dla CM"] = (
            merged_reports_3["Sales Quantity"]
            * merged_reports_3["Selling price: jedn. cena sprzedaży"]
        )
        merged_reports_3["wartość zakupu dla CM"] = (
            merged_reports_3["Sales Quantity"]
            * merged_reports_3["Purchase price: jedn. cena zakupu"]
        )
        merged_reports_3["wartość sprzedaży z uwzl. CM"] = (
            merged_reports_3["PLN"] + merged_reports_3["wartość sprzedaży dla CM"]
        )
        merged_reports_3[
            "marża (wartość sprzedaży z uwzgl. CM minus wartość zakupu dla CM)"
        ] = (
            merged_reports_3["wartość sprzedaży z uwzl. CM"]
            - merged_reports_3["wartość zakupu dla CM"]
        )

        merged_reports_3.drop_duplicates(
            subset=None, keep="first", inplace=True, ignore_index=True
        )

        merged_reports_3.to_excel(f"Credit_Memo_{credit_memo_number}.xlsx", index=False)

        merged_reports_3_test_invoice = pd.DataFrame()
        merged_reports_3_test_invoice["Invoice_reports"] = merged_reports_3["Invoice"]
        credit_memo_test_invoice = credit_memo["Invoice"]

        merged_test = pd.merge(
            merged_reports_3_test_invoice,
            credit_memo_test_invoice,
            left_on=["Invoice_reports"],
            right_on=["Invoice"],
            how="outer",
        )

        merged_test_filtered = merged_test[merged_test.isna().any(axis=1)]
        print(merged_test_filtered)

        merged_test_filtered.to_excel(f"Test_{credit_memo_number}.xlsx")

        # =========================================================================

        # marged_reports = pd.merge(
        #     sales_raport,
        #     credit_memo,
        #     left_on=["Invoice Number", "SanDisk Part Number"],
        #     right_on=["Invoice", "Debit MPN"],
        # )

        # marged_reports = marged_reports[show_columns]

        # marged_reports.rename(columns=rename_columns, inplace=True)

        # marged_reports["USD"] = (
        #     marged_reports["Sales Quantity"] * marged_reports["rabat"]
        # )
        # marged_reports["Kurs USD"] = [
        #     float(str(i).replace(",", ".")) for i in marged_reports["Kurs USD"]
        # ]
        # marged_reports["PLN"] = marged_reports["USD"] * marged_reports["Kurs USD"]
        # marged_reports["wartość sprzedaży dla CM"] = (
        #     marged_reports["Sales Quantity"]
        #     * marged_reports["Selling price: jedn. cena sprzedaży"]
        # )
        # marged_reports["wartość zakupu dla CM"] = (
        #     marged_reports["Sales Quantity"]
        #     * marged_reports["Purchase price: jedn. cena zakupu"]
        # )
        # marged_reports["wartość sprzedaży z uwzl. CM"] = (
        #     marged_reports["PLN"] + marged_reports["wartość sprzedaży dla CM"]
        # )
        # marged_reports[
        #     "marża (wartość sprzedaży z uwzgl. CM minus wartość zakupu dla CM)"
        # ] = (
        #     marged_reports["wartość sprzedaży z uwzl. CM"]
        #     - marged_reports["wartość zakupu dla CM"]
        # )

        # marged_reports.to_excel("test.xlsx", index=False)

        # print(marged_reports.info())


read_credit_memo = ReadCreditMemoRaport()
read_sales = ReadSalesRaport()


credit_memo = read_credit_memo.read_credit_memo_raport()
sales_raport = read_sales.read_sales_raport()

join_raports = JoinRaports(credit_memo, sales_raport)
join_raports.join_raports()


# print(credit_memo, sales_raport)
