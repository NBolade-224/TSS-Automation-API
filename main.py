import requests, os, json
import pandas as pd
import tkinter.filedialog as fd
from base64 import b64encode
from TssDocReferences import DocReferences
from datetime import datetime

class TSSAutomation:
    def __init__(self) -> None:
        #Passed_TSS_File = pd.read_excel(fd.askopenfilename(title='Choose TSS Export file'))
        Passed_TSS_File = pd.read_excel("\\FileName.xlsx")
        Passed_TSS_Data_Sum = Passed_TSS_File.groupby(["SupDec","PO Number"], as_index=False)["Item Price / Amount"].sum()
        Passed_DDI_File = self.consilidateExcelData()
        Passed_DDI_File_Sum = Passed_DDI_File.groupby(["PO_Number","Postcode"], as_index=False)["Item_Invoice_Amount"].sum()
        Passed_DDI_File_Sum = Passed_DDI_File_Sum.rename(columns={"PO_Number": "PO Number"})

        self.result = pd.merge(Passed_TSS_Data_Sum, Passed_DDI_File_Sum, how="left", on="PO Number")
        self.townJson = json.load(open('IrelandTowns.json'))

        self.PandasDict = {"SupDec":[],"Reason":[]}

        usern = 'username'
        passw = 'password'
        userAndPass = b64encode(b"%s:%s" % (bytes(usern,  encoding='utf-8'),bytes(passw,  encoding='utf-8'))).decode("ascii")
        self.ses = requests.Session()

        self.hrs = {
            'Accept':'application/json',
            'Content-Type':'application/json',
            'Request':'application/json',
            'Authorization' : 'Basic %s' %  userAndPass
            }

        self.Endpoint = "https://api.tradersupportservice.co.uk/api/x_fhmrc_tss_api/v1/tss_api/"

    def ExportData(self, StatusFilter: str):
 
        # FINDS ALL DRAFT SUPS AND TURNS THEM INTO A LIST OF IDS
        DraftSUPs = self.ses.get("https://api.tradersupportservice.co.uk/api/x_fhmrc_tss_api/v1/tss_api/supplementary_declarations?filter=status=%s" % StatusFilter,headers=self.hrs)                     ##### THIS IS A LIST OF ALL SUP REFERENCES
        SUPDecList = DraftSUPs.json()['result'] ## list of all sups

        # INTIAILISE PANDAS/EXCEL COLUMN HEADERS
        PandasDict = {"SupDec":[],"PO_Number":[],"Goodsid":[],"Goods":[],"PackageMark":[],"Commodity Code":[],"Item Gross Mass (KG)":[],"Item Price / Amount":[],"Number of Packages":[],"Arrival Date":[]}

        # LOOP THROUGH ALL SUP ID'S
        for current_iter, each_sup in enumerate(SUPDecList):
            print(str(each_sup['number'])+" "+str(current_iter+1)+"/"+str(len(SUPDecList)))
            print()
            List_Of_All_Goods_In_A_Sup = "https://api.tradersupportservice.co.uk/api/x_fhmrc_tss_api/v1/tss_api/goods?sup_dec_number=%s" % each_sup['number']['goods']
            Sup_details = "https://api.tradersupportservice.co.uk/api/x_fhmrc_tss_api/v1/tss_api/supplementary_declarations?reference=%s&fields=arrival_date_time,status,trader_reference,movement_reference_number" % each_sup['number']

            # LOOP THROUGH ALL GOODS FOR EACH SUP ID
            for each_good in List_Of_All_Goods_In_A_Sup:
                Good_Details = "https://api.tradersupportservice.co.uk/api/x_fhmrc_tss_api/v1/tss_api/goods?reference=%s&fields=gross_mass_kg,commodity_code,item_invoice_amount,number_of_packages,package_marks,goods_description" % each_good['goods_id'] #&fields=gross_weight_kg,commodity_code,item_invoice_amount,additional_information
                PandasDict['SupDec'].append(each_sup['number'])
                PandasDict['Goodsid'].append(each_good['goods_id'])
                PandasDict['Goods'].append(Good_Details['goods_description'])
                PandasDict['PackageMark'].append(Good_Details['package_marks'])
                PandasDict['Commodity Code'].append(Good_Details['commodity_code'])
                PandasDict['Item Gross Mass (KG)'].append(Good_Details['gross_mass_kg'])
                PandasDict['Item Price / Amount'].append(Good_Details['item_invoice_amount'])
                PandasDict['Number of Packages'].append(Good_Details['number_of_packages'])
                PandasDict['Arrival Date'].append(Sup_details['arrival_date_time']) 
                PandasDict['PO Number'].append(Sup_details['trader_reference']) 
        
        # EXPORT ALL DATA TO PANDAS
        df = pd.DataFrame(PandasDict)
        df.to_excel("./TSS Data Export %s %s.xlsx" % (StatusFilter,datetime.today().strftime('%d%m%y')))

    def consilidateExcelData(self):
        all_items_in_folder = fd.askopenfilenames(title='select all Manifest files')
        list_of_data_frames = []
        new_data_frame = pd.DataFrame()

        for item in all_items_in_folder:
            if item[-5:] == ".xlsx":
                list_of_data_frames.append(pd.read_excel(item))

        for dataframe in list_of_data_frames:
            print(dataframe)
            new_data_frame = new_data_frame.append(dataframe)

        return new_data_frame

    def priceCheck(self, objSup: object) -> None:
        if objSup["Item Price / Amount"] == objSup["Item_Invoice_Amount"]:
            self.getTown(objSup)
        else:
            print("TSS and Manifest Prices not matched")
            self.addErrorToExcel(objSup['SupDec'],"TSS and Manifest Prices not matched")
            return

    def getTown(self, objSup: object) -> None:
        try:
            Postcode = objSup['Postcode'].replace(' ', '') # remove spaces to find true len
            if Postcode[:2] == "BT" and len(Postcode) == 7:
                searchCode = Postcode[:4]
            else:
                searchCode = Postcode[:3]
    
            town = self.townJson["dictOfTowns"][searchCode]
            self.update_sup_header(objSup['SupDec'],town) 
        except:
            print("Town Error")
            self.addErrorToExcel(objSup['SupDec'],"Town not found")
            return
    
    def update_sup_header(self, Sup: str, delivery_town: str) -> None:
        json_current_details = self.ses.get(self.Endpoint+"supplementary_declarations?reference=%s&fields=declaration_choice,controlled_goods,additional_procedure,\
        goods_domestic_status,exporter_eori,arrival_date_time,total_packages,movement_type,carrier_eori,nationality_of_transport,identity_no_of_transport,\
        postponed_vat,incoterm,delivery_location_country,delivery_location_town" % Sup, headers=self.hrs)
        current_details = json_current_details.json()["result"]

        payload = {
        "op_type":"update",
        "sup_dec_number":Sup,
        "declaration_choice":"H1", 
        "controlled_goods":"no",
        "additional_procedure":"no",
        "goods_domestic_status":current_details["goods_domestic_status"],
        "exporter_eori":current_details["exporter_eori"],
        "total_packages":current_details["total_packages"],
        "movement_type":current_details["movement_type"],
        "nationality_of_transport":current_details["nationality_of_transport"],
        "identity_no_of_transport":current_details["identity_no_of_transport"],
        "postponed_vat":current_details["postponed_vat"],
        "freight_charge_currency":"GBP",
        "insurance_currency":"GBP",
        "vat_adjust_currency":"GBP",
        "incoterm":"DDP",
        "delivery_location_country":current_details["delivery_location_country"],
        "delivery_location_town":delivery_town
        }

        response = self.ses.post(self.Endpoint+"supplementary_declarations",json=payload, headers=self.hrs)
        if str(response) == "<Response [200]>":
            print("success header "+str(Sup))
            self.update_sup_goods(Sup)    
        else:
            print(str(Sup)+" ERROR")
            self.addErrorToExcel(Sup, response.json())

    def update_sup_goods(self, Sup: str) -> None:
        Goods_Json = self.ses.get(self.Endpoint+"goods?sup_dec_number=%s" % Sup,headers=self.hrs)
        list_of_goods = Goods_Json.json()['result']['goods'] ## LIST OF ALL GOODS IN A SUP DEC

        for goods in list_of_goods:
            json_current_details = self.ses.get(self.Endpoint+"goods?reference=%s&fields=type_of_packages,number_of_packages,package_marks,\
            gross_mass_kg,net_mass_kg,goods_description,invoice_number,preference,commodity_code,country_of_origin,item_invoice_amount,\
            item_invoice_currency,procedure_code,additional_procedure_code,valuation_method,valuation_indicator,nature_of_transaction,\
            payable_tax_currency,ni_additional_information_codes,country_of_preferential_origin,document_references" % goods['goods_id'], headers=self.hrs)
            current_details = json_current_details.json()["result"]

            payload = {
            "op_type":"update",
            "goods_id":goods['goods_id'],
            "type_of_packages":current_details['type_of_packages'],
            "number_of_packages":current_details['number_of_packages'],
            "package_marks":current_details['package_marks'],
            "gross_mass_kg":current_details['gross_mass_kg'],
            "net_mass_kg":current_details['gross_mass_kg'], 
            "goods_description":current_details['goods_description'],
            "invoice_number":current_details['invoice_number'],
            "preference":"300", 
            "commodity_code":current_details['commodity_code'],
            "country_of_origin":"GB", 
            "country_of_preferential_origin":"GB", 
            "item_invoice_amount":current_details['item_invoice_amount'],
            "item_invoice_currency":"GBP",
            "procedure_code":"4000", 
            "additional_procedure_code":current_details['additional_procedure_code'],
            "valuation_method":current_details['valuation_method'],
            "valuation_indicator":current_details['valuation_indicator'],
            "nature_of_transaction":current_details['nature_of_transaction'],
            "payable_tax_currency":current_details['payable_tax_currency'],
            "ni_additional_information_codes":"TCA"
            }

            if int(current_details['commodity_code']) == 9403500000:
                payload["document_references"] = [DocReferences.U110,DocReferences.Y900,DocReferences.Y904]

            elif int(current_details['commodity_code']) == 9404299000 or int(current_details['commodity_code']) == 9404291000 or int(current_details['commodity_code']) == 9404909000:
                List_of_doc_codes = []
                for codes in json_current_details.json()["result"]["document_references"]:
                    List_of_doc_codes.append(codes["document_code"])
                if "Y922" in List_of_doc_codes:
                    payload["document_references"] = [DocReferences.U110,
                    {
                    "op_type": "update",
                    "document_code": "Y922", 
                    "document_reference": "no cat or dog fur",
                    "document_reason": "no cat or dog fur"}
                    ]
                else:
                    payload["document_references"] = [DocReferences.U110, DocReferences.Y922]        
            
            elif int(current_details['commodity_code']) == 4911990000 or int(current_details['commodity_code']) == 7604291090:
                payload["document_references"] = [DocReferences.U110,DocReferences.Y069]
            
            elif int(current_details['commodity_code']) == 4001290000:
                payload["document_references"] = [DocReferences.U110,DocReferences.Y923]

            elif int(current_details['commodity_code']) == 6217100090:
                payload["document_references"] = [DocReferences.U110,DocReferences.Y904,DocReferences.Y900,DocReferences.Y922,DocReferences.Y032]

            else:
                payload["document_references"] = [DocReferences.U110]
            
            response = self.ses.post(self.Endpoint+"goods",json=payload, headers=self.hrs)
            if str(response) == "<Response [200]>":
                print("success goods "+str(Sup))
                pass  
            else:
                print(str(Sup)+" ERROR")
                self.addErrorToExcel(Sup, response.json())
                return
        
        self.submit_declation(Sup)

    def submit_declation(self, Sup: str) -> None:
        payload = {
        "op_type":"submit",
        "sup_dec_number": Sup
        }
        response = self.ses.post(self.Endpoint+"supplementary_declarations",json=payload, headers=self.hrs)
        if str(response) == "<Response [200]>":
            print("success submit "+str(Sup))
        else:
            print(str(Sup)+" ERROR")
            self.addErrorToExcel(Sup, response.json())

    def addErrorToExcel(self, Sup: str, Reason: str) -> None:
        self.PandasDict['SupDec'].append(Sup)
        self.PandasDict['Reason'].append(Reason)

    def Main(self) -> None:
        for index, objSup in self.result.iterrows():
            print("%d/%d" % (index+1,len(self.result)))
            self.priceCheck(objSup)
        
        df = pd.DataFrame(self.PandasDict)
        print(len(df))
        print(df)
        if len(df) > 0:
            df.to_excel("./Combined Manifest Data %s.xlsx" % datetime.today().strftime('%d%m%y'))

TSSAutomation().Main()