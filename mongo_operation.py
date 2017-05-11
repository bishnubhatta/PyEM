# Along with CDC2 of change maintain a log of all changes eg: ObjectID+What is the change item+what is the change value
# Update from CSV should also follow CDC2 approach
# Emailing functionality
# CI calculation
# SEAT_NUMBER Reporting
# Supervisor Change Reporting
# Polling

# Inserting data into mongodb from CSV
# mongoimport -d db_name -c collection_name --type tsv --file data.tsv -f id,url,visits,unique_visits --numInsertionWorkers 8


class mongo_oper:
    em_connect = ''
    lcr_connect = ''
    rate_connect = ''
    misc_connect = ''

    def __init__(self):
        try:
            from pymongo import MongoClient
            import urllib
            # This client is for local mongodb installation
            client = MongoClient('localhost:27017')
            #password = urllib.quote_plus('@Pr0fessi0nal')
            #username = urllib.quote_plus('bishnubhatta')
            # This client is for cloud version of mongodb installation
            #client = MongoClient('mongodb://'+ username + ':' + password + '@ds137370.mlab.com:37370')
            my_db = client.pyem_local
            self.em_connect = my_db.EM_RESOURCES
            self.lcr_connect = my_db.LCR_DETAILS
            self.rate_connect = my_db.RATE_CARD
            self.misc_connect = my_db.MISC_INFO
            print 'Connection opened successfully!!!'
        except Exception, e4:
            print str(e4)

    # This function is not yet used
    def mongoimport_data_dump_from_file(self, db_name):
        try:
            from pymongo import MongoClient

            file_path=raw_input('\n Please enter the tsv file path \n')
            file_name=raw_input('\n Please enter the tsv file name \n')
            db_name = raw_input('\n Please enter the db name \n ')
            collection_name = raw_input('\n Please enter the collection name \n')

            import_string = 'mongoimport -h localhost:27017 -p 27017 -d ' + db_name + ' -c ' + collection_name + ' --type tsv ' + ' --file ' + file_path + '\\' + file_name
            MongoClient(import_string)
            print 'Import successful!!! Please verify the loaded data'
        except Exception, e4:
            print str(e4)

    # This function is not yet used
    def update_lcr_from_csv_by_sapid(self, csv_file_name, update_field_name):
        try:
            import csv, pymongo, sys
            csvreader = csv.reader(open(csv_file_name, 'rb'), delimiter=',', quotechar='"')
            for line in csvreader:
                empid, updt_field_value = line
                self.lcr_connect.update({"SAP_ID": int(empid)},{"$set": {str(update_field_name): str(updt_field_value)}})
        except Exception, e:
            print str(e)

    # This function is not yet used
    def active_document_single_field_update(self,crieteriafldnm, crieteriafldtyp, criteria_value, updtfldnm, updtfldval):
        try:
            # from datetime import datetime
            from bson.objectid import ObjectId
            updt_criteria = int(criteria_value) if crieteriafldtyp == 'int' else str(criteria_value)
            # get count
            count = self.em_connect.find({
                "$and": [{str(crieteriafldnm): updt_criteria}, {"RECORD_ACTIVE_INDICATOR": "Y"}]}).count()
            #print "Active records present step1: "+ str(count)
            # if multiple active rows found abort
            if count != 1:
                print "\n Multiple or Zero Active records found for the record."
                return -1
            else:
                # get the active row into orig
                orig=self.em_connect.find({
                    "$and": [{str(crieteriafldnm): updt_criteria}, {"RECORD_ACTIVE_INDICATOR": "Y"}]})
                # Insert the saved copy (orig) of old record
                for item in orig:
                    item["_id"]=ObjectId()
                    # print (str(item["RECORD_ACTIVE_INDICATOR"]))
                    print item
                    self.em_connect.insert(item)
                # update RECORD_ACTIVE_INDICATOR to "N"
                self.em_connect.update_one({
                    "$and":[{str(crieteriafldnm): updt_criteria},{"RECORD_ACTIVE_INDICATOR": "Y"}]},
                    {'$set': {"RECORD_ACTIVE_INDICATOR": "N"}})
                # Verify that only 1 active record is present
                new_count = self.em_connect.find({
                    "$and": [{str(crieteriafldnm): updt_criteria},
                             {"RECORD_ACTIVE_INDICATOR": "Y"}]}).count()
                # after insertion and updation, count of active records should be 1
                if new_count <> 1:
                    print "\n Multiple or Zero Active records found for the record."
                else:
                    self.em_connect.update_one(
                        {"$and": [{str(crieteriafldnm): updt_criteria},{"RECORD_ACTIVE_INDICATOR": "Y"}]},
                            {'$set': {str(updtfldnm): str(updtfldval)}})
                    print "\nRecord updated successfully\n"
                return 0
        except Exception, e:
            print str(e)
            return -1


    # Utility function for dictionary to list conversion
    def dict_to_list(self,dict):
        bbb = [v for k, v in dict.items()]
        return bbb

    # Get distinct Enterprise ID from mongo to populate in the dropdown list
    def get_distinct_ent_id(self):
        try:
            r = []
            my_data = self.em_connect.find({"RECORD_ACTIVE_INDICATOR": "Y"},{"ENTERPRISE_ID": 1, "_id": 0})
            for data in my_data:
                r.append(data["ENTERPRISE_ID"])
            return [] if my_data is None else r
        except Exception, e1:
            print str(e1)

    # Function to get distinct WO for GUI drop down
    def get_distinct_wo(self):
        try:
            r=[]
            my_data = self.em_connect.find({},{"WORK_ORDER_NUMBER": 1,"_id": 0}).distinct("WORK_ORDER_NUMBER")
            return [] if my_data is None else my_data
        except Exception, e1:
            print str(e1)

    def get_distinct_team(self):
        try:
            r=[]
            my_data = self.em_connect.find({}, {"WO_TEAM": 1, "_id": 0}).distinct("WO_TEAM")
            return [] if my_data is None else my_data
        except Exception, e1:
            print str(e1)

    # Read the raw XLS excluding the header
    def read_em_raw_data(self):
        import xlrd
        path = 'C:/PyEM/Input/EM.xlsx'
        workbook = xlrd.open_workbook(path)

        worksheet = workbook.sheet_by_index(0)
        # Set to -1 if you want to include the only header data present in the file
        offset = 0
        rows = []
        for i, row in enumerate(range(worksheet.nrows)):
            if i <= offset:  # (Optionally) skip headers
                continue
            r = []
            for j, col in enumerate(range(worksheet.ncols)):
                r.append(worksheet.cell_value(i, j))
            rows.append(r)
        return rows

    # Read the EM xls row and select only the required columns.
    # Append to it 13 blank values as mongo has 13 extra fields those are not in EM
    def em_raw_data_to_mongo_data_mapping(self,xls_raw_rows):
        final_lst = []
        for k in range(0, len(xls_raw_rows)):
            aaa= xls_raw_rows[k]
            r=[]
            for l in range(0,len(aaa)):
                if l in [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,
                         34,35,36,37,38,39,40,41,42,43,47,48,49,50,51,53,54,55,56,57,58,59,60,61,62,72,77,80]:
                    # MODIFIED_DT is the 9th column of the xls and is a datetime. Convert that to string for comparision
                    # Remaining dates in the xls should also be converted.
                    if l == 8:
                        r.append(self.excel_time_to_string(aaa[l]))
                    elif l in [13,15,16,20,21,22,23,24,39,40,47,48,54,55,58,59,60,]:
                        r.append(self.excel_date_to_string(0 if aaa[l] is None else aaa[l]))
                    elif l == 0:
                        r.append(int(aaa[l]))
                    else:
                        r.append(aaa[l])
            r.extend(['' for i in range(0,15)])
            row_dict = self.list_to_dict(r)
            final_lst.append(r)
            final_lst.append(row_dict)
        return final_lst # a list of list and dict is returned eg: [[list1],{dict1},[list2],{dict2}....]

    # EM returns X number of fields, Mongo has Y number of fields. Needed a function to
    def list_to_dict(self,r):
        key_list = ["EM_ID","EMPLOYEE_ID","FIRST_NAME","LAST_NAME","ENTERPRISE_ID","WORK_ORDER_NUMBER","ROLE_MAPPING",
                    "MODIFIED_BY","MODIFIED_DT","ACTIVE_RESOURCE","BILLABLE_FLAG","BUFFER_FLAG","BENCH_FLAG",
                    "ACCENTURE_DOJ","LEVEL","EFFECTIVE_FROM","TRAVELERS_START_DT","ABACUS_PROJECT","ATTRITION_FLAG",
                    "ATTRITION_REASON","BILLABLE_END_DT","BILLABLE_START_DT","BIRTH_DT","BUFFER_END_DT",
                    "BUFFER_START_DT","CITIZENSHIP","LOB","PROJECT_MODULE_AREA","CURRENT_LOCATION",
                    "CDP_ROLL_OFF_FORM_ATTACHED","CDP_ROLL_ON_FORM_ATTACHED","EMERGENCY_CONTACT1",
                    "RELATIONSHIP1","EMERGENCY_CONTACT2","RELATIONSHIP2","GCP_EMPLOYEE_ID",
                    "GSO_ROLE","HOME_ADDRESS","HOME_PHONE","I94_END_DT","I94_START_DT","LOANED","MIDDLE_NAME",
                    "MOBILE_PHONE","ONSITE_END_DT","ONSITE_START_DT","PASSPORT_NUMBER","PRIMARY_SKILL","PROD_ID",
                    "SHADOW","SHADOW_PERIOD_END_DT","SHADOW_PERIOD_START_DT","TIC_ID","TRAVELERS_EMAIL",
                    "TRAVELERS_END_DT","VISA_END_DT","VISA_START_DT","VISA_TYPE","WO_TEAM","DU_NAME",
                    "BUFFER_CATEGORY","NEW_IT_CONVERSANT","SUPERVISOR_ENTERPRISE_ID","SUPERVISOR_ST_DT","SUPERVISOR_END_DT",
                    "RECORD_ACTIVE_INDICATOR","CERTIFICATIONS","NJ_ORIENTATION_STS","CDP_COMPLETION_STS",
                    "IFT_COMPLETION_STS","SEAT_NUMBER","RRD_TYPE","VM_NAME","HL_DATE",
                    "BENCH_BUFFER_AGING","POLLING1","POLLING2"]
        final_dict = dict(zip(key_list,r))
        return final_dict

    # Excel reads dates as numbers. Use this to convert to string formatted datetime for comparision
    def excel_time_to_string(self,xltimeinput):
        try:
            import xlrd
            import datetime
            retVal = xlrd.xldate.xldate_as_datetime(xltimeinput, 0)
            retVal=retVal.strftime('%m/%d/%Y %H:%M')
        except ValueError:
            retVal = xltimeinput
        return retVal

    # Excel reads dates as numbers. Use this to convert to string formatted date for loading
    def excel_date_to_string(self,xltimeinput):
        try:
            import xlrd
            import datetime
            retVal = xlrd.xldate.xldate_as_datetime(xltimeinput, 0)
            retVal=retVal.strftime('%m/%d/%Y')
        except ValueError:
            retVal = ''
        return retVal

    # Send the Mongo returned datetime to get string formatted datetime for comparision
    def python_time_to_string(self,timeinput):
        try:
            from datetime import datetime
            retVal = datetime.strptime(timeinput, '%m/%d/%Y %H:%M')
            retVal=retVal.strftime('%m/%d/%Y %H:%M')
        except ValueError:
            retVal = timeinput
        return retVal

    # Compare the MODIFIED_BY and MODIFIED_DT fields to identify any change
    def compare_em_mongo_data(self, em_rows):
        import xlrd
        import datetime
        try:
            # em_rows looks like [[list1],{dict1},[list2],{dict2}....]. so index=[even]=>list and index=[odd]=>dict
            final_dict = []
            for i in range(0,len(em_rows),2):
                 mongo_list = []
                 temp_em_dict_rows = em_rows[i+1] #take the dictionary
                 temp_em_rows = em_rows[i]
                 ret_dict = {} if self.fetch_active_record_by_em_id(int(temp_em_rows[0])) is None else self.fetch_active_record_by_em_id(int(temp_em_rows[0]))
                 # Outstanding Issue : EM_ID not present EM but present in mongo will appear in reports.
                 if len(ret_dict) > 0:
                    mongo_list.append(ret_dict["MODIFIED_BY"])
                    mongo_list.append(self.python_time_to_string(ret_dict["MODIFIED_DT"]))
                    if (str(temp_em_rows[7]) != str(mongo_list[0]) or (str(temp_em_rows[8]) != str(mongo_list[1]))):
                        final_dict.append(temp_em_dict_rows)
                        #print 'CHANGED RECORD'
                    else:
                        print 'NO_CHANGE RECORD'
                 else:
                    final_dict.append(temp_em_dict_rows)
                    #print 'NEW RECORD'
            return final_dict
        except Exception, e1:
             print str(e1)

    # Fields present in Mongo for current record but not in latest EM record, are updated in this function
    def replace_non_em_fields_from_old_record(self,final_dict):
        try:
            return_dict = []
            for i in range(0,len(final_dict)):
                temp_rec = final_dict[i]
                my_data = [] if self.em_connect.find_one({"$and": [{"EM_ID": int(temp_rec["EM_ID"])},
                                                             {"RECORD_ACTIVE_INDICATOR": "Y"}]}
                    , {"SUPERVISOR_ENTERPRISE_ID":1,"SUPERVISOR_ST_DT":1,"SUPERVISOR_END_DT":1,
                    "RECORD_ACTIVE_INDICATOR":1,"CERTIFICATIONS":1,"NJ_ORIENTATION_STS":1,
                    "CDP_COMPLETION_STS":1,"IFT_COMPLETION_STS":1,"SEAT_NUMBER":1,"RRD_TYPE":1,"VM_NAME":1,
                    "HL_DATE":1,"BENCH_BUFFER_AGING": 1,"POLLING1":1,"POLLING2":1, "_id": 0}) is None else self.em_connect.find_one({"$and": [{"EM_ID": int(temp_rec["EM_ID"])},
                                                             {"RECORD_ACTIVE_INDICATOR": "Y"}]}
                    , {"SUPERVISOR_ENTERPRISE_ID":1,"SUPERVISOR_ST_DT":1,"SUPERVISOR_END_DT":1,
                    "RECORD_ACTIVE_INDICATOR":1,"CERTIFICATIONS":1,"NJ_ORIENTATION_STS":1,
                    "CDP_COMPLETION_STS":1,"IFT_COMPLETION_STS":1,"SEAT_NUMBER":1,"RRD_TYPE":1,"VM_NAME":1,
                    "HL_DATE":1,"BENCH_BUFFER_AGING": 1,"POLLING1":1,"POLLING2":1, "_id": 0})
                if len(my_data) > 0:
                    for item in temp_rec:
                        temp_rec["SUPERVISOR_ENTERPRISE_ID"] = my_data["SUPERVISOR_ENTERPRISE_ID"]
                        temp_rec["SUPERVISOR_ST_DT"] = my_data["SUPERVISOR_ST_DT"]
                        temp_rec["SUPERVISOR_END_DT"] = my_data["SUPERVISOR_END_DT"]
                        temp_rec["RECORD_ACTIVE_INDICATOR"] = "Y"
                        temp_rec["CERTIFICATIONS"] = my_data["CERTIFICATIONS"]
                        temp_rec["NJ_ORIENTATION_STS"] = my_data["NJ_ORIENTATION_STS"]
                        temp_rec["CDP_COMPLETION_STS"] = my_data["CDP_COMPLETION_STS"]
                        temp_rec["IFT_COMPLETION_STS"] = my_data["IFT_COMPLETION_STS"]
                        temp_rec["SEAT_NUMBER"] = my_data["SEAT_NUMBER"]
                        temp_rec["RRD_TYPE"] = my_data["RRD_TYPE"]
                        temp_rec["VM_NAME"] = my_data["VM_NAME"]
                        temp_rec["HL_DATE"] = my_data["HL_DATE"]
                        temp_rec["BENCH_BUFFER_AGING"] = my_data["BENCH_BUFFER_AGING"]
                        temp_rec["POLLING1"] = my_data["POLLING1"]
                        temp_rec["POLLING2"] = my_data["POLLING2"]
                else:
                    for item in temp_rec:
                        temp_rec["RECORD_ACTIVE_INDICATOR"] = "Y"
                return_dict.append(temp_rec)
            return return_dict
        except Exception, e1:
            print str(e1)

    # You now enter the rows from previous function into Mongo using insertOne
    # Before that, Expire the corresponding old row using updateOne
    def insert_em_rows_into_mongo(self,insert_rows):
        try:
            temp_stor = []
            if len(insert_rows) > 0:
                for i in range(0,len(insert_rows)):
                    temp_stor = insert_rows[i]
                    # See if the record already exists in mongo if yes then update old + insert new else insert new
                    if len(self.fetch_active_record_by_em_id(temp_stor["EM_ID"])) > 0:
                        # Get the ACTIVE records (ideally there should be only 1) and update the RECORD_ACTIVE_INDICATOR to N
                        self.em_connect.update_one(
                            {"$and": [{"EM_ID": int(temp_stor["EM_ID"])}, {"RECORD_ACTIVE_INDICATOR": "Y"}]},
                            {"$set": {"RECORD_ACTIVE_INDICATOR": "N"}})
                        self.em_connect.insert(temp_stor)
                    else:
                        self.em_connect.insert(temp_stor)
        except Exception, e1:
            print str(e1)

    # Test with no change, change and new record scenario

    def fetch_active_record_by_em_id(self,my_id):
        try:
            my_data = self.em_connect.find_one({"$and": [{"EM_ID": my_id},{"RECORD_ACTIVE_INDICATOR": "Y"}]}
                                    ,{"MODIFIED_BY": 1, "MODIFIED_DT": 1, "_id": 0})
            return [] if my_data is None else my_data
        except Exception, e1:
            print str(e1)

    def expire_records_not_present_in_EM(self,em_id):
        try:
            my_data = self.em_connect.find({"RECORD_ACTIVE_INDICATOR": "Y"},{"EM_ID": 1,"_id": 0}).distinct("EM_ID")
            for data in my_data:
                if len(self.fetch_active_record_by_em_id(my_data["EM_ID"])) == 0:
                    print "TB Coded"
        except Exception, e1:
            print str(e1)

    ###################################################################################################################
    def view_all_documents_for_entid(self,entid):
        try:
            my_data = self.em_connect.find({"ENTERPRISE_ID": entid},{"_id":0})
            self.print_html_report(my_data,str(entid)+"_employee_data")
        except Exception, e1:
            print str(e1)

    def view_all_documents_for_WO(self,wo_list):
        try:
            wo_data = wo_list.split(",")
            for wo in wo_data:
                my_data = self.em_connect.find({"WORK_ORDER_NUMBER": wo},{"_id":0})
                print str(my_data)
                self.print_html_report(my_data,str(wo)+"_employee_data")
        except Exception, e1:
            print str(e1)

    def logical_delete_by_entid(self, entid):
        try:
            count = self.em_connect.find(
                {"$and": [{"ENTERPRISE_ID": entid},{"RECORD_ACTIVE_INDICATOR": "Y"}]}).count()
            print count
            if count > 0:
                self.em_connect.update(
                    {"$and": [{"ENTERPRISE_ID": entid},{"RECORD_ACTIVE_INDICATOR": "Y"}]},
                    {'$set': {"RECORD_ACTIVE_INDICATOR": "N"}})
                print '\nLogical deletion successful for ENTERPRISE_ID: ' + entid + '\n'
            else:
                print '\n No records available for logical deletion for the enterprise id mentioned \n'
        except Exception, e2:
            print str(e2)

    def close_conn(self):
        self.em_connect.close()
        print '\n Connection closed successfully!!!\n'

    def fname_lname_to_entid_mapping(self, fname, lname):
        try:
            import re
            regex = ".*" + fname #+ ".*"
            my_data = self.em_connect.find_one({"$and": [{"FIRST_NAME": re.compile(regex, re.IGNORECASE)},
                                                     {"RECORD_ACTIVE_INDICATOR": "Y"}]}, {"ENTERPRISE_ID": 1, "_id": 0})
            if my_data is None:
                return ""
            else:
                return my_data['ENTERPRISE_ID']
        except Exception, e1:
            print(str(e1))

    def fetch_bill_rate(self, role, location):
        try:
            location = "ONSHORE" if location[0:2] in('Ha','Hu') \
                else ("NO_LOCATION" if location == 'NO_LOCATION' else "OFFSHORE")
            role = "NO_ROLE" if role == 'NO_ROLE' else role
            my_data = self.rate_connect.find_one({"$and": [{"Role": role},{"Location": location}]},{"_id": 0})
            return my_data['Rate']
        except Exception, e1:
            print(str(e1))

# This function fetches the role of an employeeid
    def fetch_role_for_entid(self, entid):
        try:
            my_data = self.em_connect.find_one({"$and": [{"ENTERPRISE_ID": entid}, {"RECORD_ACTIVE_INDICATOR":"Y"}]},
                                               {"ROLE_MAPPING": 1, "_id": 0})
            return "NO_ROLE" if my_data is None else my_data['ROLE_MAPPING']
        except Exception, e1:
            print(str(e1))

# This function fetches the role of an employeeid
    def fetch_location_for_entid(self, entid):
        try:
            my_data = self.em_connect.find_one({"$and":[{"ENTERPRISE_ID": entid},{"RECORD_ACTIVE_INDICATOR": "Y"}]},
                                               {"CURRENT_LOCATION": 1, "_id": 0})
            return "NO_LOCATION" if my_data is None else my_data['CURRENT_LOCATION']
        except Exception, e1:
            print(str(e1))

# This function fetches the GCP_EMP_ID if available
    def fetch_gcpid_for_entid(self, entid):
        try:
            my_data = self.em_connect.find_one({"$and": [{"ENTERPRISE_ID": entid}, {"RECORD_ACTIVE_INDICATOR": "Y"}]},
            {"GCP_EMPLOYEE_ID": 1,"EMPLOYEE_ID": 1, "CURRENT_LOCATION": 1, "_id": 0})
            retval = str(my_data["EMPLOYEE_ID"]) if str(my_data["GCP_EMPLOYEE_ID"])== '' else str(my_data["GCP_EMPLOYEE_ID"])
            return retval
        except Exception, e1:
            print(str(e1))

# This function fetches the LCR of an employeeid
    def fetch_lcr_for_employeeid(self, employeeid):
        try:
            cnt = self.lcr_connect.find({"SAP_ID": int(employeeid)}).count()
            lcr = self.lcr_connect.find_one({"SAP_ID": int(employeeid)},{"LCR": 1, "_id": 0})
            return 0 if cnt == 0 else lcr['LCR']
        except Exception, e1:
            print(str(e1))

    def get_billable_flag_for_entid(self, entid):
        try:
            myflag = self.em_connect.find_one({"$and": [{"ENTERPRISE_ID": entid}, {"RECORD_ACTIVE_INDICATOR": "Y"}]},
            {"BILLABLE_FLAG": 1, "_id": 0})
            return myflag['BILLABLE_FLAG']
        except Exception, e1:
            print str(e1)

    def fetch_wo_for_entid(self, entid):
        try:
            wo = self.em_connect.find_one({"$and": [{"ENTERPRISE_ID": entid}, {"RECORD_ACTIVE_INDICATOR": "Y"}]},
            {"WORK_ORDER_NUMBER": 1, "_id": 0})
            return wo['WORK_ORDER_NUMBER']
        except Exception, e1:
            print str(e1)


# This function calculates CI for an enterprise id
    def calculate_ci_for_employee(self, entid):
        try:
            r = [0,0]
            billable_flag = self.get_billable_flag_for_entid(entid)
            if billable_flag == 'Y':
                r = []
                employeeid = self.fetch_gcpid_for_entid(entid)
                #print str(employeeid)
                role = self.fetch_role_for_entid(entid)
                #print role
                location=self.fetch_location_for_entid(entid)
                #print location
                billrt = self.fetch_bill_rate(role, location)
                #print str(billrt)
                workorder = self.fetch_wo_for_entid(entid)
                #print workorder
                wo_type = self.misc_connect.find_one({"FIELD": "PC_WO_LIST"}, {"_id": 0})
                pc_wo_list = wo_type["VALUE"].split(",")
                data = self.misc_connect.find_one({"FIELD": "BILLING_DISCOUNT"}, {"_id": 0})
                #print pc_wo_list
                for pc_wo in pc_wo_list:
                    if pc_wo == workorder:
                        billhr = 9
                    else:
                        billhr = 8
                    dbillrt = float(billrt * (1 - float(data["VALUE"])))
                    billing = float(billrt * billhr)
                    dbilling = float(dbillrt * billhr)
                    # LCR table is based on GCP_EMP_ID for onshore resources
                    lcr = self.fetch_lcr_for_employeeid(employeeid)
                    lcrcost = float(lcr*9)
                    #print "LCR:"+ str(lcr)
                    if lcr == 0:
                        ci = 0
                        dci = 0
                    else:
                        ci = (billing - lcrcost) * 100 / float(billing)
                        dci = (dbilling - lcrcost) * 100 / float(dbilling)
                    r.append(ci)
                    r.append(dci)
            #print r
            return r
        except Exception, e1:
            print(str(e1))

# This function calculates CI for a WO
    def calculate_ci_for_wo(self, wo):
        try:
            return_val=[]
            final_str=[]
            wo_list = wo.split(",")
            calc_ci = []
            calc_dci = []
            for workorder in wo_list:
                count = self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(workorder)},
                            {"RECORD_ACTIVE_INDICATOR": "Y"},{"BENCH_FLAG":"N"},{"BUFFER_FLAG":"N"}]}).count()
                emdata =  self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(workorder)},
            {"RECORD_ACTIVE_INDICATOR": "Y"},{"BENCH_FLAG":"N"},{"BUFFER_FLAG":"N"}]},{"ENTERPRISE_ID": 1, "_id": 0})
                ci_sum = 0
                dci_sum = 0
                for data in emdata:
                    aaa = self.calculate_ci_for_employee(data["ENTERPRISE_ID"])
                    final_str.append("<tr><td>" + data["ENTERPRISE_ID"] +
                                     "</td><td>" + str(workorder) +
                                     "</td><td>" + str(aaa[0]) +
                                     "</td><td>" + str(aaa[1])+"</tr>")
                    calc_ci.append(aaa[0])
                    calc_dci.append(aaa[1])
                    ci_sum = sum(calc_ci)
                    dci_sum = sum(calc_dci)
                final_ci_wo = ci_sum/count
                final_dci_wo = dci_sum/count
                #print 'Final CI : ' + str(final_ci_wo)
                #print 'Final dCI : ' + str(final_dci_wo)
                return_val.append([workorder,final_ci_wo,final_dci_wo])
                calc_ci=[]
                calc_dci=[]
                hrd=["ENTERPRISE_ID","WORKORDER","PRE DISCOUNT CI","POST DISCOUNT CI"]
                self.print_list_report('Employee_CI_' + wo, final_str,hrd)
            return return_val
        except Exception, e2:
            print str(e2)

    def fetch_bench_report_by_wo(self, work_order):
        try:
            wo_list=work_order.split(",")
            for wo in wo_list:
                emdata = self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(wo)}, {"RECORD_ACTIVE_INDICATOR": "Y"},
                                                        {"ACTIVE_RESOURCE": "Y"}, {"BENCH_FLAG": "Y"}]},
                            {"EMPLOYEE_ID": 1,"ENTERPRISE_ID":1,"LEVEL":1, "_id": 0})
                cnt = self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(wo)}, {"RECORD_ACTIVE_INDICATOR": "Y"},
                                                        {"ACTIVE_RESOURCE": "Y"}, {"BENCH_FLAG": "Y"}]}).count()
                if cnt > 0:
                    self.print_html_report(emdata,wo + '_bench_report')
        except Exception, e1:
            print(str(e1))

    def fetch_buffer_report_by_wo(self, work_order):
        try:
            wo_list = work_order.split(",")
            for wo in wo_list:
                emdata = self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(wo)}, {"RECORD_ACTIVE_INDICATOR": "Y"},
                                                        {"ACTIVE_RESOURCE": "Y"}, {"BUFFER_FLAG": "Y"}]},
                                                         {"EMPLOYEE_ID": 1, "ENTERPRISE_ID": 1, "LEVEL": 1, "_id": 0})
                cnt = self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(wo)}, {"RECORD_ACTIVE_INDICATOR": "Y"},
                                                        {"ACTIVE_RESOURCE": "Y"}, {"BUFFER_FLAG": "Y"}]}).count()
                if cnt > 0:
                    self.print_html_report(emdata, wo + '_buffer_report')
        except Exception, e1:
            print(str(e1))

    def rrd_type_report_by_wo(self, work_order):
        try:
            wo_list = work_order.split(",")
            for wo in wo_list:
                emdata = self.em_connect.find({"$and": [{"WORK_ORDER_NUMBER": str(wo)},
                                                        {"RECORD_ACTIVE_INDICATOR": "Y"}, {"ACTIVE_RESOURCE": "Y"}]},
                                              {"EMPLOYEE_ID": 1, "ENTERPRISE_ID": 1, "LEVEL": 1, "RRD_TYPE":1,"_id": 0})
                self.print_html_report(emdata, wo + '_rrd_type')
        except Exception, e1:
            print(str(e1))

    def fetch_polling_report(self,poll):
        try:
            pn =self.misc_connect.find_one({"FIELD":poll},{"POLLING_NAME":1,"_id":0})
            polldata = self.em_connect.find({"$and": [{"RECORD_ACTIVE_INDICATOR":"Y"},{poll:"Y"}]},
                                            {"EMPLOYEE_ID": 1, "ENTERPRISE_ID": 1,"WORK_ORDER_NUMBER": 1,"_id": 0})
            self.print_html_report(polldata, pn['POLLING_NAME'] + '_polling')
        except Exception, e1:
            print(str(e1))


    def fetch_pyramid_report_by_wo(self,work_order):
        try:
            final_str = []

            wo_list = work_order.split(",")
            for wo in wo_list:
                num_of_resource = 0
                aaa = self.em_connect.aggregate([
                    {'$match': {'$and': [{"WORK_ORDER_NUMBER": str(wo)},{"RECORD_ACTIVE_INDICATOR": "Y"}, {"ACTIVE_RESOURCE": "Y"}]}},
                    {'$group': {'_id' : "$LEVEL" , 'COUNT': {'$sum': 1}}}])
                num_of_resource = self.em_connect.find({'$and': [{"WORK_ORDER_NUMBER": str(wo)},{"RECORD_ACTIVE_INDICATOR": "Y"}, {"ACTIVE_RESOURCE": "Y"}]}).count()
                final_data = []
                # Here we calculate the % of all the level
                for data in aaa:
                    idc_tgt = self.misc_connect.find_one({"$and": [{"FIELD": "IDC_TARGET_PYRAMID"},{"LEVEL": data["_id"]}]}, {"_id": 0})
                    final_str.append("<tr><td>" + data["_id"] +
                                     "</td><td>" + str(idc_tgt["VALUE"]) +
                                     "</td><td>" + str(data["COUNT"]) +
                                     "</td><td>" + str(float(data["COUNT"])*100/num_of_resource) + "</tr>")
                hrd = ["Level", "IDC Target", "Current Count", "Current Percentage"]
                self.print_list_report('Pyramid_Report_' + wo, final_str, hrd)
        except Exception, e1:
            print(str(e1))

    # As of now this function is not used anywhere
    def print_dictionary_value(self, dictionary):
        try:
            for x, y in dictionary.items():
                print str(x) + ":" + str(y)
            print '\n'
        except Exception, e1:
            print(str(e1))

    def print_html_report(self, cursor, rpt_name):
        try:
            import HTML
            wo = rpt_name.split("_")[0]
            rpt_type = rpt_name.split("_")[1].capitalize()

            HTMLFILE = 'C:/PyEM/Output/' + rpt_name + '.html'
            f = open(HTMLFILE, 'w')
            header = []
            bodydata = []
            i = 0
            j = 0
            for data in cursor:
                sorted(data)
                if i == 0:
                    for myrow in data:
                        bodydata.append(data[myrow].encode('utf8') if isinstance(data[myrow],basestring) == True else unicode(data[myrow]).encode('utf8'))
                        header.append(myrow)
                        t = HTML.Table(header_row=header)
                        i += 1
                else:
                    for myrow in data:
                        bodydata.append(
                            data[myrow].encode('utf8') if isinstance(data[myrow], basestring) == True else unicode(
                                data[myrow]).encode('utf8'))
            aaa = [bodydata[len(header) * i:len(header) * (i + 1)] for i in range(len(bodydata)/len(header))]
            for data in aaa:
                t.rows.append(data)
            htmlcode = str(t)
            f.write('<html><body><h1>')
            f.write(rpt_type + " Report For :"+ wo)
            f.write('</h1>')
            f.write('<p>')
            f.write(htmlcode)
            f.write('<p>')
        except Exception, e1:
            print str(e1)
#################################################################
# This function was written specifically for printing CI data
    def print_list_report(self,rpt_name,rpt_data,hdr_data):
        try:
            HTMLFILE = 'C:/PyEM/Output/' + rpt_name + '.html'
            f = open(HTMLFILE, 'w')
            f.write('<html><body><h1>')
            f.write(rpt_name + " Report ")
            f.write('</h1>')
            f.write('<table border = "1">')
            f.write('<tr>')
            for head in hdr_data:
                f.write('<th>')
                f.write(head)
                f.write('</th>')
            f.write('</tr>')
            for data in rpt_data:
                f.write(data)
        except Exception, e1:
            print str(e1)


