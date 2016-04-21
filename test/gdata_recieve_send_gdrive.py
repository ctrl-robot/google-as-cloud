#!/usr/bin/python

#import gspread
import datetime
import sys
import time
import serial
import gdata.spreadsheet.service
import gdata.docs.service
import urllib
import string

class Worksheet(object):
    """Worksheet wrapper class.
    """
    def __init__(self, gd_client, spreadsheet_key, worksheet_key):
        """Initialise a client

        :param gd_client:
            A GDATA client.
        :param spreadsheet_key:
            A string representing a google spreadsheet key.
        :param worksheet_key:
            A string representing a google worksheet key.
        """
        self.gd_client = gd_client
        self.spreadsheet_key = spreadsheet_key
        self.worksheet_key = worksheet_key
        self.keys = {'key': spreadsheet_key, 'wksht_id': worksheet_key}
        self.entries = None

    def flush_entry_cache(self):
        self.entries = None

    def _row_to_dict(self, row):
        """Turn a row of values into a dictionary.
        :param row:
            A row element.
        :return:
            A dict.
        """
        return dict([(key, row.custom[key].text) for key in row.custom])

    def _get_row_entries(self):
        """Get Row Entries

        :return:
            A rows entry
        """
        if not self.entries:
            self.entries = self.gd_client.GetListFeed(**self.keys).entry
        return self.entries

    def get_rows(self):
        """Get Rows

        :return:
            A list of rows dictionaries.
        """
        return [self._row_to_dict(row) for row in self._get_row_entries()]

    def update_row(self, index, row_data):
        """Update Row
        :param index:
            An integer designating the index of a row to update (zero based).
        :param row_data:
            A dictionary containing row data.
        :return:
            A row dictionary for the updated row.
        """
        entries = self._get_row_entries()
        rows = self.get_rows()
        rows[index].update(row_data)
        entry = self.gd_client.UpdateRow(entries[index], rows[index])
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row update failed: '{0}'".format(entry))
        self.entries[index] = entry
        return self._row_to_dict(entry)

    def insert_row(self, row_data):
        """Insert Row

        :param row_data:
            A dictionary containing row data.
        :return:
            A row dictionary for the inserted row.
        """
        entry = self.gd_client.InsertRow(row_data, **self.keys)
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row insert failed: '{0}'".format(entry))
        if self.entries:
            self.entries.append(entry)
        return self._row_to_dict(entry)

    def delete_row(self, index):
        """Delete Row

        :param index:
            A row index.
        """
        entries = self._get_row_entries()
        self.gd_client.DeleteRow(entries[index])
        if self.entries:
            del self.entries[index]

    def delete_all_rows(self):
        """Delete All Rows
        """
        entries = self._get_row_entries()
        for entry in entries:
            self.delete_row(entry)


def get_items(feed):
    """ Get the items in the feed.
     
    Either a list of documents that the user has or a list of worksheets within
    a given spreadsheet.
 
    Keyword arguments:
    feed -- The feed (xml file).
    """
    items = {}
    for entry in feed.entry:
        id_parts = urllib.unquote(entry.id.text).replace(':','/').split('/')
        key = id_parts[len(id_parts) - 1]
        items[entry.title.text.lower()] = key
    return items



try:
    sname='AC Router System'
    wname='Sheet1'
    dname=time.strftime('data_%Y%m%d')
    gname='graph'
    category = ''
    email = 'yourgmailaccount@gmail.com'
    password = 'yourpassword'
    source = category 
    spreadsheet_name = sname.lower()
    worksheet_name = wname.lower()
    datasheet_name = dname.lower()
    graphsheet_name = gname.lower()
except NameError:
    cname = ''



gd_client = gdata.docs.service.DocsService()
gd_client.ClientLogin(email, password, source=source)

client = gdata.spreadsheet.service.SpreadsheetsService()
client.debug = False
client.ssl = True
client.email = email
client.password = password
client.source = 'test client'
client.ProgrammaticLogin()


if not category:
    feed = gd_client.GetDocumentListFeed()
else:
    query = gdata.docs.service.DocumentQuery(categories=[category])
    feed = gd_client.Query(query.ToUri())

docs = get_items(feed)
curr_key = docs[spreadsheet_name]

    # Get the WorkSheet within the SpreadSheet.
feed = client.GetWorksheetsFeed(curr_key)
 
sheets = get_items(feed)
    #print sheets # Comment this out to see the worksheets a given spreadsheet has.

curr_wksht_id = sheets[worksheet_name]
graph_wksht_id = sheets[graphsheet_name]

try:
    data_wksht_id = sheets[datasheet_name]
except: #create new sheet if the required sheet doesn't exist
    client.AddWorksheet(datasheet_name,1,13,curr_key)
    feed = client.GetWorksheetsFeed(curr_key)
    sheets = get_items(feed)
    data_wksht_id = sheets[datasheet_name]
    #write header for the new sheet
    entry = client.UpdateCell(1, 1, 'machine',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 2, 'date',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 3, 'time',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 4, 'm1',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 5, 'm3',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 6, 'm4',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 7, 'm5',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 8, 'm6',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 9, 'm7',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 10, 'm9',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 11, 'm10',curr_key,data_wksht_id)
    entry = client.UpdateCell(1, 12, 'm11',curr_key,data_wksht_id)
    #entry = client.InsertRow(1, 13, 'machine',curr_key,data_wksht_id)
    
    if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
        print "Create new sheet succeeded."
    else:
        print "Create new sheet failed."

curr_wksht=Worksheet(client,curr_key,curr_wksht_id)

curr_wksht.gd_client=client
curr_wksht.spreadsheet_key=curr_key
curr_wksht.worksheet_key=curr_wksht_id


    
#ser = serial.Serial(10, 38400, timeout=0.1,parity=serial.PARITY_NONE, rtscts=1)
#s = ser.read(100)       # read up to one hundred bytes
                        # or as much is in the buffer
#line = ser.readline()   # read a '\n' terminated line
                        # write a string  ser.write("hello")

data_num=0

while True:
    #send command to pc1,pc2,pc3 to get data
    #ser.write("PC1")        # CALL PC1 TO RETRIVE DATA
    #s = ser.readline(100)       # read up to one hundred bytes
    s="PC3 10 12 25 65 89 54 74 36 65 14"
    if len(s)>0:
        datastr=s.split();
        if datastr[0]=="PC3":# and datastr[9]=='\n':
            gdict = {}
            gdict['date'] = time.strftime('%m/%d/%Y')
            gdict['time'] = time.strftime('%m/%d/%Y %H:%M:%S')
            gdict['machine'] = datastr[0]
            gdict['m1'] = datastr[1]
            gdict['m3'] = datastr[2]
            gdict['m4'] = datastr[3]
            gdict['m5'] = datastr[4]
            gdict['m6'] = datastr[5]
            gdict['m7'] = datastr[6]
            gdict['m9'] = datastr[7]
            gdict['m10'] = datastr[8]
            gdict['m11'] = datastr[9]

            # save history data for future analysis
            entry = client.InsertRow(gdict, curr_key, data_wksht_id)
            if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
               data_num=data_num+1#print "Insert row to history succeeded."
            else:
                print "Insert row to history failed."

            # prepare data for graph
            entry = client.InsertRow(gdict, curr_key, graph_wksht_id)
            if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
               data_num=data_num+1#print "Insert row to graph succeeded."
            else:
                print "Insert row to graph failed."
                
            # save data for realtime display
            try:
                entry = curr_wksht.insert_row(0,gdict)
                #data_num=data_num+1
                #print "Data updated at: "+gdict['time']+" No."+str(data_num)
                #print "Update row of realtime display succeeded."
            except:
                print "Update row of realtime display failed."              

    time.sleep(1)
    
