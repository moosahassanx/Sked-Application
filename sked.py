# Imports
from datetime import datetime
from fileinput import filename
from io import BytesIO
import streamlit as st
from datetime import datetime
from msilib.schema import tables
import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlrd
from openpyxl import Workbook, load_workbook
import xlsxwriter

def show_sked_page():
    st.title('Sked Software')

    uploaded_file = st.file_uploader('Upload BV Shipping List')
    if uploaded_file is not None:
        # Read file as object
        shippingSheet = pd.read_excel(uploaded_file, sheet_name=None)

        # Retrieving data
        URL = "https://nccports.portauthoritynsw.com.au/eports/mobilemovements.asp"
        page = requests.get(URL)
        soup = BeautifulSoup(page.content, "html.parser")

        # Setting out
        vesselDates = soup.find_all('table', attrs={'border':'0'})
        vesselValues = soup.find_all('table', attrs={'border':'1'})
        vesselValues.pop(0)
        vessels = []
        vesselNames = []
        tempVesselNames = []

        # Parse dates
        for i in vesselDates:
            date = i.find("b").string
            vessels.append({"date": date, "data": []})

        # Parse data
        for vesselIndex, j in enumerate(vesselValues):
            cells = j.findChildren("font")
            cellIndex = 0
            rowData = []
            for cell in cells:
                rowData.append(cell.string.strip())
                cellIndex += 1
                if(cellIndex == 5):
                    vessels[vesselIndex]["data"].append({
                        "Time": rowData[0],
                        "From": rowData[1],
                        "To": rowData[2],
                        "Vessel": rowData[3],
                        "Loa": rowData[4]
                    })
                    tempVesselNames.append(rowData[3])
                    cellIndex = 0
                    rowData = []


        st.text('============= nccports.portauthoritynsw.com.au reading =================')
        st.text(vessels)
        st.text(tempVesselNames)

        # TODO: parse CSV data to draw comparisons between Mobile Movements and the BV Shipping List (ask rizwan bhai how it works)
        st.text("============= Reading BV Shipping List VESSELS =================")
        shippingSheet = pd.read_excel("BV Shipping List 2022 - Email.xlsm", sheet_name=None)
        shippingSheetName, df = next(iter(shippingSheet.items()))
        df.columns = df.iloc[1]
        df = df[2:]

        # column cleanup
        df = df[df['VESSEL'].notna()]

        colVessel = df["VESSEL"]
        for vessel in colVessel:
            st.text(vessel)

        # Comparison between vessels and colVessel
        for day in vessels:
            for vesselName in day['data']:
                if(vesselName['Vessel'] in tempVesselNames):
                    vesselNames.append(vesselName['Vessel'])

        # TODO: create new xlsm as target
        st.text("============= PLACEHOLDER SHEET =================")
        # wb = Workbook()
        # ws = wb.active
        # ws['A1'] = 42
        # ws.append(vesselNames)
        # wb.save('new_document.xlsm')
        # wb = load_workbook('new_document.xlsm', keep_vba=True)
        # wb.save('new_document.xlsm')

        # output = BytesIO()
        # workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        # worksheet = workbook.add_worksheet()

        # worksheet.write('A1', 'Hello')
        # workbook.close()

        st.write('============================= FILE DOWNLOADING  ======================')
        # https://xlsxwriter.readthedocs.io/
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        fileName = datetime.today().strftime('%Y-%m-%d') + ' Sheet1.xlsx'

        worksheet.write('A1', 'Hello')
        workbook.close()

        st.download_button(
            label="Download Sheet1",
            data=output.getvalue(),
            file_name=fileName,
            mime="application/vnd.ms-excel"
        )

        