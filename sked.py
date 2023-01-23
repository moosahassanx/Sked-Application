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
import io
import base64

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


        st.header('NCCPorts Shipping Movements Data Reading')
        st.table(vessels)
        st.header('Extracting vessel names')
        st.table(tempVesselNames)

        # create the excel file
        df = pd.DataFrame(tempVesselNames)
        filename = 'my_objects.xlsx'
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()

        @st.cache(allow_output_mutation=True)
        def serve_excel():
            return open(filename, "rb").read()

        if st.button("Download Excel File"):
            st.write("Downloading...")
            b = serve_excel()
            b64 = base64.b64encode(b).decode()  # some strings <-> bytes conversions necessary here
            href = f'<a href="data:file/xlsx;base64,{b64}">Download</a>'
            st.markdown(href, unsafe_allow_html=True)
                