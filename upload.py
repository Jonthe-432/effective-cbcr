from flask import Flask, request, render_template, Response
import pandas as pd
import xml.etree.ElementTree as ET
import datetime

x = datetime.datetime.now()

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        xlsfile = request.files.get('file')
        MessageSpecDF = pd.read_excel(xlsfile, sheet_name=0, skiprows=1, usecols=[1,2], index_col=0)
        ReportingEntityDF = pd.read_excel(xlsfile, sheet_name=1, skiprows=1, usecols=[1,2], index_col=0)
        Table1DF = pd.read_excel(xlsfile, sheet_name=2, skiprows=2, index_col=0)
        Table2DF = pd.read_excel(xlsfile, sheet_name="TABLE2_Constituent_Entities", skiprows=3, usecols=[i for i in range(1,32)], index_col=0)
        Table3DF = pd.read_excel(xlsfile, sheet_name="TABLE3_Additional Info", skiprows=5, usecols=[i for i in range(1,7)])

        Table1DF = Table1DF.dropna()
        Table1DF = Table1DF.fillna(0)
        Table3DF = Table3DF[Table3DF['Other information'].notna()]

        #set xml root element
        top = ET.Element("cbc:CBC_NL")

        #set namespaces
        iso = top.set("xmlns:iso", "urn:belastingdienst:ISOtypes:3.0.V1")
        cbc = top.set("xmlns:cbc", "http://xml.belastingdienst.nl/schemas/CBCNL/3.0/01")
        stf = top.set("xmlns:stf", "urn:belastingdienst:BDtypes:3.0.V1")
        xsi = top.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        version = top.set("version", "3.0/01")
        schemaLocation = top.set("xsi:schemaLocation", "http://xml.belastingdienst.nl/schemas/CBCNL/3.0/01 CBCNL_3.0_V1.20181109.xsd")

        FileType = ET.SubElement(top, "cbc:FileType")
        FileType.text = "CBCNLGEG"

        #CBC_NL/MessageSpec/
        MessageSpec = ET.SubElement(top, "cbc:MessageSpec")

        SendingEntityIN = ET.SubElement(MessageSpec, "cbc:SendingEntityIN")
        SendingEntityIN.text = str(MessageSpecDF.iloc[4,0])

        SendingEntityName = ET.SubElement(MessageSpec, "cbc:SendingEntityName")
        SendingEntityName.text = ReportingEntityDF.iloc[3,0]

        TransmittingCountry = ET.SubElement(MessageSpec, "cbc:TransmittingCountry")
        TransmittingCountry.text = "NL"

        MessageType = ET.SubElement(MessageSpec, "cbc:MessageType")
        MessageType.text = "CBC"

        MessageRefId = ET.SubElement(MessageSpec, "cbc:MessageRefId")
        MessageRefId.text = TransmittingCountry.text + str((datetime.datetime.now().year - 1))+"-" + SendingEntityIN.text

        ReportingPeriodBegin = ET.SubElement(MessageSpec, "cbc:ReportingPeriodBegin")
        ReportingPeriodBegin.text = str(x.year - 1) + "-01-01"


        ReportingPeriod = ET.SubElement(MessageSpec, "cbc:ReportingPeriod")
        ReportingPeriod.text = str(x.year - 1) + "-12-31"

        Timestamp = ET.SubElement(MessageSpec, "cbc:Timestamp")
        Timestamp.text = str(x.strftime("%Y-%m-%dT%H:%M:%S"))


        #CBC_NL/CbcBody/

        CbcBody = ET.SubElement(top, "cbc:CbcBody")

        ReportingEntity = ET.SubElement(CbcBody, "cbc:ReportingEntity")

        #CBC_NL/CbcBody/ReportingEntity

        Entity = ET.SubElement(ReportingEntity, "cbc:Entity")

        #CBC_NL/CbcBody/ReportingEntity/Entity

        ResCountryCode = ET.SubElement(Entity, "cbc:ResCountryCode")
        ResCountryCode.text = "NL"

        TIN = ET.SubElement(Entity, "cbc:TIN")
        TIN.text = str(ReportingEntityDF.iloc[5,0])
        TIN.set("issuedBy", "NL")

        Name = ET.SubElement(Entity, "cbc:Name")
        Name.text = str(ReportingEntityDF.iloc[3,0])

        ReportingEntity_Address = ET.SubElement(Entity, "cbc:Address")

        #CBC_NL/CbcBody/ReportingEntity/Entity/Address

        ReportingEntityCountryCode = ET.SubElement(ReportingEntity_Address, "cbc:CountryCode")
        ReportingEntityCountryCode.text = "NL"

        ReportingEntityAddressFix = ET.SubElement(ReportingEntity_Address, "cbc:AddressFix")

        #CBC_NL/CbcBody/ReportingEntity/Entity/Address/AddressFix

        ReportingEntityAddressStreet = ET.SubElement(ReportingEntityAddressFix, "cbc:Street")
        ReportingEntityAddressStreet.text = str(ReportingEntityDF.iloc[23,0])

        ReportingEntityAddressBuildingIdentifier = ET.SubElement(ReportingEntityAddressFix, "cbc:BuildingIdentifier")
        ReportingEntityAddressBuildingIdentifier.text = str(ReportingEntityDF.iloc[24,0])

        ReportingEntityAddressCity = ET.SubElement(ReportingEntityAddressFix, "cbc:City")
        ReportingEntityAddressCity.text = str(ReportingEntityDF.iloc[30,0])

        ReportingRoleDict = {"Ultimate Parent Entity": "CBC701" ,  "Surrogate parent Entity" : "CBC702", "Local Filing": "CBC703"}


        ReportingRole = ET.SubElement(ReportingEntity, "cbc:ReportingRole" )
        ReportingRole.text = ReportingRoleDict[ReportingEntityDF.iloc[10,0]]

        ReportingEntityDocSpec = ET.SubElement(ReportingEntity, "cbc:DocSpec" )

        ReportingEntityDocTypeIndic = ET.SubElement(ReportingEntityDocSpec,"stf:DocTypeIndic")
        ReportingEntityDocTypeIndic.text = "OECD1"

        ReportingEntityDocRefId = ET.SubElement(ReportingEntityDocSpec,"stf:DocRefId")
        ReportingEntityDocRefId.text = ResCountryCode.text + str(x.year - 1) + "-" + str(ReportingEntityDF.iloc[5,0]) + str(x.strftime("%Y%m%d%H%M%S"))


        DocTypeIndicDict = {"New Data": "OECD1" ,  "Corrected Data" : "OECD1", "Deletion of Data": "OECD3"}

        #CBC_NL/CbcBody/CbcReports

        for index, row in Table1DF.iterrows():
            CbcReports = ET.SubElement(CbcBody, "cbc:CbcReports")

            CbcReportsDocSpec = ET.SubElement(CbcReports, "cbc:DocSpec")
            CbcReportsDocTypeIndic = ET.SubElement(CbcReportsDocSpec, "stf:DocTypeIndic")
            CbcReportsDocTypeIndic.text = DocTypeIndicDict[ReportingEntityDF.iloc[1,0]]
            CbcReportsDocRefId = ET.SubElement(CbcReportsDocSpec, "stf:DocRefId")
            CbcReportsDocRefId.text = "NL" + str(x.year - 1) + "-" + str(ReportingEntityDF.iloc[5,0]) + str(x.strftime("%Y%m%d%H%M%S"))

            ResCountryCode = ET.SubElement(CbcReports, "cbc:ResCountryCode")
            ResCountryCode.text = index

            Summary = ET.SubElement(CbcReports, "cbc:Summary")

            Revenues = ET.SubElement(Summary, "cbc:Revenues")

            Unrelated = ET.SubElement(Revenues, "cbc:Unrelated")
            Unrelated.set("currCode", MessageSpecDF.iloc[7,0])
            Unrelated.text = str(round(row['Unrelated party revenue']))

            Related = ET.SubElement(Revenues, "cbc:Related")
            Related.set("currCode", MessageSpecDF.iloc[7,0])
            Related.text = str(round(row['Related party revenue']))

            Total = ET.SubElement(Revenues, "cbc:Total")
            Total.set("currCode", MessageSpecDF.iloc[7,0])
            Total.text = str(round(row['Total Revenue']))

            ProfitOrLoss = ET.SubElement(Summary, "cbc:ProfitOrLoss")
            ProfitOrLoss.set("currCode", MessageSpecDF.iloc[7,0])
            ProfitOrLoss.text = str(round(row['Profit (loss) before income tax']))

            TaxPaid = ET.SubElement(Summary, "cbc:TaxPaid")
            TaxPaid.set("currCode", MessageSpecDF.iloc[7,0])
            TaxPaid.text = str(round(row['Income tax paid (on a cash basis)']))

            TaxAccrued = ET.SubElement(Summary, "cbc:TaxAccrued")
            TaxAccrued.set("currCode", MessageSpecDF.iloc[7,0])
            TaxAccrued.text = str(round(row['Income Tax accrued - Current year']))

            Capital = ET.SubElement(Summary, "cbc:Capital")
            Capital.set("currCode", MessageSpecDF.iloc[7,0])
            Capital.text = str(round(row['Stated capital']))

            Earnings = ET.SubElement(Summary, "cbc:Earnings")
            Earnings.set("currCode", MessageSpecDF.iloc[7,0])
            Earnings.text = str(round(row['Accumulated earnings']))

            NbEmployees = ET.SubElement(Summary, "cbc:NbEmployees")
            NbEmployees.text = str(round(row['Number of employees']))

            Assets = ET.SubElement(Summary, "cbc:Assets")
            Assets.set("currCode", MessageSpecDF.iloc[7,0])
            Assets.text = str(round(row['Tangible assets other than Cash and Cash Equivalents']))

        #ConstEntities

            for index2, row2 in Table2DF.loc[[index]].iterrows():

                ConstEntities = ET.SubElement(CbcReports, "cbc:ConstEntities")

                ConstEntity = ET.SubElement(ConstEntities, "cbc:ConstEntity")

                ResCountryCode = ET.SubElement(ConstEntity, "cbc:ResCountryCode")
                ResCountryCode.text = row2['Country Code']

                TIN = ET.SubElement(ConstEntity, "cbc:TIN")
                TIN.text = str(row2['Tax Identification Number (TIN)'])
                TIN.set("issuedBy", row2['Country Code'])


                Name = ET.SubElement(ConstEntity, "cbc:Name")
                Name.text = row2['Constituent Entities resident in the Tax Jurisdiction']

                Address = ET.SubElement(ConstEntity, "cbc:Address")

                if pd.notnull(row2['Country Code']):
                    CountryCode = ET.SubElement(Address, "cbc:CountryCode")
                    CountryCode.text = row2['Country Code']

                AddressFix = ET.SubElement(Address, "cbc:AddressFix")



                if pd.notnull(row2['Street']):
                    Street = ET.SubElement(AddressFix, "cbc:Street")
                    Street.text = row2['Street']


                if pd.notnull(row2['Building']):
                    BuildingIdentifier = ET.SubElement(AddressFix, "cbc:BuildingIdentifier")
                    BuildingIdentifier.text = str(row2['Building'])

                if pd.notnull(row2['Suite']):
                    SuiteIdentifier = ET.SubElement(AddressFix, "cbc:SuiteIdentifier")
                    SuiteIdentifier.text = str(row2['Suite'])

                if pd.notnull(row2['Floor']):
                    FloorIdentifier = ET.SubElement(AddressFix, "cbc:FloorIdentifier")
                    FloorIdentifier.text = str(row2['Floor'])

                if pd.notnull(row2['District']):
                    DistrictName = ET.SubElement(AddressFix, "cbc:DistrictName")
                    DistrictName.text = row2['District']

                if pd.notnull(row2['PO Box']):
                    POB = ET.SubElement(AddressFix, "cbc:POB")
                    POB.text = str(row2['PO Box'])

                if pd.notnull(row2['Post Code']):
                    PostCode = ET.SubElement(AddressFix, "cbc:PostCode")
                    PostCode.text = str(row2['Post Code'])

                if pd.notnull(row2['City']):
                    City = ET.SubElement(AddressFix, "cbc:City")
                    City.text = row2['City']

                if pd.notnull(row2['Country Subentity']):
                    CountrySubentity = ET.SubElement(AddressFix, "cbc:CountrySubentity")
                    CountrySubentity.text = row2['Country Subentity']

                if pd.notnull(row2['Tax Jurisdiction of organisation or incorporation if different from Tax Jurisdiction of Residence']):
                    IncorpCountryCode = ET.SubElement(ConstEntities, "cbc:IncorpCountryCode")
                    IncorpCountryCode.text = row2['Tax Jurisdiction of organisation or incorporation if different from Tax Jurisdiction of Residence']


                if pd.notnull(row2['Research and Development']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC501"

                if pd.notnull(row2['Holding/managing intellectual property']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC502"

                if pd.notnull(row2['Purchasing or Procurement']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC503"

                if pd.notnull(row2['Manufacturing or Production']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC504"

                if pd.notnull(row2['Sales, Marketing or Distribution']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC505"

                if pd.notnull(row2['Administrative, Management or Support Services']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC506"

                if pd.notnull(row2['Provision of services to unrelated parties']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC507"

                if pd.notnull(row2['Internal Group Finance']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC508"

                if pd.notnull(row2['Regulated Financial Services']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC509"

                if pd.notnull(row2['Insurance']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC510"

                if pd.notnull(row2['Holding shares or other equity instruments']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC511"

                if pd.notnull(row2['Dormant']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC512"

                if pd.notnull(row2['Other']):
                    BizActivities = ET.SubElement(ConstEntities, "cbc:BizActivities")
                    BizActivities.text = "CBC513"
                    OtherEntityInfo = ET.SubElement(ConstEntities, "cbc:OtherEntityInfo")
                    OtherEntityInfo.text = row2['If "Other" was selected, please specify']


        #CBC_NL/CbcBody/AdditionalInfo


        if len(Table3DF) > 0:

            for index, row in Table3DF.iterrows():
                AdditionalInfo = ET.SubElement(CbcBody, "cbc:AdditionalInfo")
                DocSpec = ET.SubElement(AdditionalInfo, "cbc:DocSpec")

                DocTypeIndic = ET.SubElement(DocSpec, "stf:DocTypeIndic")
                DocTypeIndic.text = DocTypeIndicDict[row['Document Type Indicator']]

                DocRefId = ET.SubElement(DocSpec, "stf:DocRefId")
                DocRefId.text = "NL" + str(x.year - 1) + "-" + str(ReportingEntityDF.iloc[5,0]) + str(x.strftime("%Y%m%d%H%M%S"))

                OtherInfo = ET.SubElement(AdditionalInfo, "cbc:OtherInfo")
                OtherInfo.text = row['Other information']

        xml = ET.tostring(top)

        return Response(xml,
                   mimetype="text/xml",
                   headers={"Content-Disposition":
                                "attachment;filename=cbcr.xml"})
    return render_template('/upload.html')

if __name__ == '__main__':
    app.run(debug=True)
