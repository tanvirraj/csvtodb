__author__ = 'tanvir'

from datetime import datetime

import xlrd


book = xlrd.open_workbook("Industry Canada List.xlsx")

print book.sheet_names()

first_sheet = book.sheet_by_index(0)

# print "read rows...."
# for i in range(4, 100):
#     print first_sheet.row_values(i)

print "start inserting db ..."




def enter_into_db(data):
    import psycopg2
    from psycopg2.extensions import AsIs

    agency_list = []


    try:
        conn = psycopg2.connect("dbname='jobdb' user='jobmin' password='f1d3r!@#' host='localhost'")
        print "Database Connected"
    except:
        print "I am unable to connect to the database."

    cur = conn.cursor()
    print data
    for i in range(2, 507):
        o = first_sheet.row_values(i)
        print o 
        # try:
        print "$$$$$$$$$$$$$$$$$$$"
        print "%s,%s" % (o[24], o[25])
        print "$$$$$$$$$$$$$$$$$$$"
        name = o[1][0:200]
        print " type of ....."
        print type(o[2])
        if o[2] is float:
            phone = str(o[2])
        else:
            phone = o[2]
        if o[3] is float:
            toll_free_phone = str(o[3])
        else:                
            toll_free_phone = o[3][0:100]
        if o[4] is float:
            crisis_phone = str(o[4])
        else:    
            crisis_phone = o[4]
        if o[5] is float:
            fax = str(o[5])
        else:
            fax = o[5]    
        email = o[6][0:150]
        website = o[7][0:150]
        address_1 = o[8][0:250]
        address_2 = o[9][0:250]
        city = o[10][0:250]
        province = o[11][0:250]
        postal_code = o[12][0:250]
        # location = o[13]
        if o[14] is float:
            fees = str(o[14])
        else:
            fees = o[14]
        if o[15] is float:
            hours = str(o[15])
        else:
            hours = o[15]
        language_of_service = o[17][0:300]
        organization_type = o[18][0:300]
        eligibility = o[19][0:200]
        how_to_apply = o[20][0:500]
        physical_access = o[21][0:300]
        about = o[22][0:500]
        tag = o[23][0:500]
        service_description_title = o[24][0:500]
        service_description = o[25][0:500]
        contact = o[26][0:500]
        if o[27] != "":
            lat = o[27]
        else:
            lat = 0.0 

        if o[28] != "":
            lon = o[28]
        else:
            lon = 0.0 
       
        fb_url = o[29][0:500][0:500]
        twitter = o[30][0:500]
        linkedin =o[31][0:500][0:500]
        approved = True
        print o 
#         cur.execute(
# ...     """INSERT INTO some_table (an_int, a_date, a_string)
# ...         VALUES (%s, %s, %s);""",
# ...     (10, datetime.date(2005, 11, 18), "O'Reilly"))
        
        cur.execute(
            """INSERT INTO agency_agency (name,phone,toll_free_phone,crisis_phone,fax,email,website,
            mail_address,street_address,city,province,postal_code,fees,hours,
            language_of_service,organization_type,eligibility,how_to_apply,
            physical_access,about,tag,service_description_title, service_description,contact,
            lat,lon, fb_url, twitter, linkedin, approved)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id;""", 
            (name,phone,toll_free_phone,crisis_phone,fax,email,website,
            address_1,address_2,city,province,postal_code,fees,hours,
            language_of_service,organization_type,eligibility,how_to_apply,
            physical_access,about,tag,service_description_title, service_description,contact,
            lat,lon, fb_url, twitter, linkedin, approved))
        print "Inserted to Agency"
        agency = cur.fetchone()[0]
        print agency
        agency_list.append(agency)
        print "agency list ...."
        print agency_list
        conn.commit()



    service_type_list = ['Business Networks', 'Professional Association', 'Career Counselling']
    # service_type_list = ['Education & Certifications', 'Campus', 'Career Counseling']
    for name in service_type_list:
        cur.execute(
        """INSERT INTO agency_servicetype (name)
        VALUES (%s)RETURNING id; """, (name,))
        servicetype_id = cur.fetchone()[0]
        conn.commit()
        
        for k in agency_list:
            cur.execute(
            """INSERT INTO agency_servicetype_agencies(servicetype_id, agency_id)
            VALUES (%s,%s); """, (servicetype_id,k,))
            conn.commit()

        print "inserted service type"



enter_into_db(first_sheet)
