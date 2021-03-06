import pandas as pd
from dateutil import tz
from lxml import etree


def create_xml(user_input_file, user_input_result):

    df = pd.read_excel(user_input_file)
    first_col_name = df.columns[1]
    xmlns_uris = {'xsi': "http://www.w3.org/2001/XMLSchema-instance",
                  None: "http://www.pse.pl/sire/ws/Common"}

    root = etree.Element("PlannedResourceSchedule", nsmap=xmlns_uris)
    type = etree.SubElement(root, "type")
    type.text = df.iloc[0][first_col_name]

    time_interval = etree.SubElement(root, "schedule_Period.timeInterval")

    start_date = etree.SubElement(time_interval, "start")
    start_date.text = df.iloc[1][first_col_name].astimezone(
        tz.UTC).strftime("%Y-%m-%dT%H:%MZ")
    end_date = etree.SubElement(time_interval, "end")
    end_date.text = df.iloc[2][first_col_name].astimezone(
        tz.UTC).strftime("%Y-%m-%dT%H:%MZ")
    n = 0
    for l in df.columns:
        if l == 'Opis':
            continue
        n += 1
        PlannedResource = etree.SubElement(root, "PlannedResource_TimeSeries")

        mRID = etree.SubElement(PlannedResource, "mRID")
        mRID.text = str(n)
        businessType = etree.SubElement(PlannedResource, "businessType")
        businessType.text = df.iloc[3][l]
        measurement_Unit = etree.SubElement(
            PlannedResource, "measurement_Unit.name")
        measurement_Unit.text = df.iloc[4][l]
        registeredResource = etree.SubElement(
            PlannedResource, "registeredResource.mRID")
        registeredResource.text = df.iloc[5][l]

        Series_Period = etree.SubElement(PlannedResource, "Series_Period")

        date = etree.SubElement(Series_Period, "timeInterval")
        st_date = etree.SubElement(date, "start")
        st_date.text = df.iloc[1][l].astimezone(
            tz.UTC).strftime("%Y-%m-%dT%H:%MZ")
        e_date = etree.SubElement(date, "end")
        e_date.text = df.iloc[2][l].astimezone(
            tz.UTC).strftime("%Y-%m-%dT%H:%MZ")
        res = etree.SubElement(Series_Period, "resolution")
        res.text = df.iloc[8][l]

        for index, row in df.iterrows():
            if index > 8:
                point = etree.SubElement(Series_Period, "Point")
                position = etree.SubElement(point, "position")
                position.text = str(index-8)
                quantity = etree.SubElement(point, "quantity")
                quantity.text = str(row[l])

    tree = etree.ElementTree(root)
    with open(user_input_result+'.xml', "wb") as files:
        tree.write(files, encoding='utf-8', xml_declaration=True)


user_input_file = input('Excel file name: ')
user_input_result = input('XML file name: ')
create_xml(user_input_file, user_input_result)
