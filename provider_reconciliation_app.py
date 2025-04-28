import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


st.set_page_config(page_title="NUITEE -Provider- Reconciliation Tool Ay", layout="centered")
st.title("🧮 SOA & Allocation Reconciliation Tool For Slaves")
st.write("Upload your files below and click **Start Reconciliation** to generate your reports.")


st.header("📂 Upload Files")

soa_file = st.file_uploader("Upload SOA File (.xlsx)", type=["xlsx"])
allocations_file = st.file_uploader("Upload Provider Allocations File (.xlsx)", type=["xlsx"])
nrb_file = st.file_uploader("Upload Not Requested Bookings File (Optional, .xlsx)", type=["xlsx"])



@st.cache_data
def get_reconciliation_output(SOA, ALLOCATIONS, merged, disputes, refunds_and_NRB, SOA_TSP_final):
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        SOA.to_excel(writer, sheet_name="SOA", index=False)
        ALLOCATIONS.to_excel(writer, sheet_name="Allocations", index=False)
        merged.to_excel(writer, sheet_name="Reconciliation", index=False)
        disputes.to_excel(writer, sheet_name="Disputes", index=False)
        refunds_and_NRB.to_excel(writer, sheet_name="Refunds & Not Requested Bkgs", index=False)
        SOA_TSP_final.to_excel(writer, sheet_name="SOA-TSP", index=False)
    output_excel.seek(0)
    return output_excel

@st.cache_data
def get_template_output(SOA_Analysis_Template):
    template_excel = BytesIO()
    with pd.ExcelWriter(template_excel, engine="xlsxwriter") as writer:
        SOA_Analysis_Template.to_excel(writer, sheet_name="SOA - To upload", index=False)
    template_excel.seek(0)
    return template_excel





if soa_file and allocations_file:
    output_filename = st.text_input("Enter Reconciliation Output File Name:", value="Reconciliation_Output.xlsx")
    template_filename = st.text_input("Enter Template Output File Name:", value="Template_Output.xlsx")

    if st.button("🔍 Start Reconciliation"):

        with st.spinner("Processing files... Please wait ⏳"):


            SOA = pd.read_excel(soa_file, header=3, sheet_name="SOA - TSP")
            ALLOCATIONS = pd.read_excel(allocations_file, header=0)

            if nrb_file:
                NRB = pd.read_excel(nrb_file, header=0, sheet_name="Rfunds & Not Requested Bkgs", usecols="A:O")
                NRB.columns = NRB.columns.str.strip()
                NRB["Nuitee Booking Id"] = NRB["Nuitee Booking Id"].astype(str).str.strip()
                NRB["Nuitee Booking Id"] = pd.to_numeric(NRB["Nuitee Booking Id"], errors="coerce")
            else:
                NRB = pd.DataFrame(columns=[
                    "Nuitee Booking Id", "Provider", "Provider Booking Id", "Reservation Date", "Hotel Name", "City Name",
                    "Country Name", "CheckIn", "CheckOut", "Holder Name", "Provider Reservation Status",
                    "ZohoInvoiceConversionRate", "ZohoBillConversionRate", "soa_amount", "CurrencyFrom"
                ])

            SOA.columns = SOA.columns.str.strip()
            ALLOCATIONS.columns = ALLOCATIONS.columns.str.strip()

            SOA["Nuitee Booking Id"] = SOA["Nuitee Booking Id"].astype(str).str.strip()
            ALLOCATIONS["A"] = ALLOCATIONS["A"].astype(str).str.strip()

            SOA["Nuitee Booking Id"] = pd.to_numeric(SOA["Nuitee Booking Id"], errors="coerce")
            ALLOCATIONS["A"] = pd.to_numeric(ALLOCATIONS["A"], errors="coerce")


            soa_grouped = SOA.groupby("Nuitee Booking Id", as_index=False).agg({
                "Provider": "first",
                "Provider Booking Id": "first",
                "Reservation Date": "first",
                "Hotel Name": "first",
                "City Name": "first",
                "Country Name": "first",
                "CheckIn": "first",
                "CheckOut": "first",
                "Holder Name": "first",
                "Provider Reservation Status": "first",
                "ZohoInvoiceConversionRate": "first",
                "ZohoBillConversionRate": "first",
                "AmountToPayToProviderCurrencyFrom": "sum",
                "CurrencyFrom": "first"
            })

            allocations_grouped = ALLOCATIONS.groupby("A", as_index=False).agg({
                "B": "sum",
                "C": "first"
            })

            soa_grouped.rename(columns={"AmountToPayToProviderCurrencyFrom": "soa_amount"}, inplace=True)
            allocations_grouped.rename(columns={
                "A": "Nuitee Booking Id",
                "B": "provider_amount",
                "C": "Currency"
            }, inplace=True)

            combined_SOA_NRB = pd.concat([soa_grouped, NRB], ignore_index=True)

            merged = pd.merge(combined_SOA_NRB, allocations_grouped, on="Nuitee Booking Id", how="outer")

            merged["SOA_Null_Values"] = np.select(
                [merged["soa_amount"].isna(), merged["soa_amount"] == 0],
                ["Not found in SOA", "Null"],
                default="-"
            )

            merged["Allocations_Null_Values"] = np.select(
                [merged["provider_amount"].isna(), merged["provider_amount"] == 0],
                ["Not found in allocations", "Null"],
                default="-"
            )

            merged["soa_amount"] = merged["soa_amount"].fillna(0)
            merged["provider_amount"] = merged["provider_amount"].fillna(0)

            merged["Difference"] = merged["soa_amount"] - merged["provider_amount"]

            disputes = merged[merged["Difference"] <= -1].copy()
            disputes["Dispute Type"] = disputes["soa_amount"].apply(lambda x: "disputed refund" if x < 0 else "disputed Booking")

            refunds_and_NRB = merged[merged["Difference"] >= 1].copy()
            refunds_and_NRB["Type"] = refunds_and_NRB["provider_amount"].apply(lambda x: "unearned refund" if x < 0 else "not requested Booking")

            soa_tsp_filter = merged.copy()

            soa_tsp_filter["Unearned Refund"] = np.where(
                (soa_tsp_filter["provider_amount"] < -1) & (soa_tsp_filter["Difference"] > 0),
                "Yes",
                "-"
            )

            soa_tsp_filter.loc[
                soa_tsp_filter["soa_amount"] > soa_tsp_filter["provider_amount"],
                "soa_amount"
            ] = soa_tsp_filter["provider_amount"]

            final_columns_order = [
                "Nuitee Booking Id", "Provider", "Provider Booking Id",
                "Reservation Date", "Hotel Name", "City Name", "Country Name",
                "CheckIn", "CheckOut", "Holder Name", "Provider Reservation Status",
                "ZohoInvoiceConversionRate", "ZohoBillConversionRate",
                "soa_amount", "CurrencyFrom", "provider_amount", "Difference", "Unearned Refund"
            ]

            SOA_TSP_final = soa_tsp_filter[final_columns_order]

            soa_Template_ToUplaod = merged[
                (merged["soa_amount"] != 0) & (merged["provider_amount"] != 0)
            ].copy()

            soa_Template_ToUplaod["SOA_Amount_Final"] = soa_Template_ToUplaod.apply(
                lambda row: row["soa_amount"] if row["CurrencyFrom"] == "EUR" or row["CurrencyFrom"] == "USD" else row["soa_amount"] * row["ZohoBillConversionRate"],
                 axis=1
            )
            
            soa_Template_ToUplaod["Currency_From_Final"] = soa_Template_ToUplaod.apply(
                lambda row: row["CurrencyFrom"] if row["CurrencyFrom"] == "EUR" or row["CurrencyFrom"] == "USD" else "EUR",
                 axis=1
            )
        

            soa_Template_ToUplaod["Conversion_Rate"] = soa_Template_ToUplaod.apply(
                lambda row: row["ZohoBillConversionRate"] if row["CurrencyFrom"] == "USD" else row["ZohoInvoiceConversionRate"],
                axis=1
            )

            final_columns_for_Analysis = [
                "Nuitee Booking Id", "Provider", "Provider Booking Id",
                "Reservation Date", "Hotel Name", "City Name", "Country Name",
                "CheckIn", "CheckOut", "Holder Name", "Provider Reservation Status",
                "Conversion_Rate",
                "SOA_Amount_Final", "Currency_From_Final"
            ]

            SOA_Analysis_Template = soa_Template_ToUplaod[final_columns_for_Analysis]

            st.session_state.output_excel = get_reconciliation_output(SOA, ALLOCATIONS, merged, disputes, refunds_and_NRB, SOA_TSP_final)
            st.session_state.template_excel = get_template_output(SOA_Analysis_Template)

            st.success("✅ Reconciliation Completed Successfully!")

    if "output_excel" in st.session_state and "template_excel" in st.session_state:
        st.download_button(
            label="📥 Download Reconciliation Report",
            data=st.session_state.output_excel,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📥 Download Template Report",
            data=st.session_state.template_excel,
            file_name=template_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )            
