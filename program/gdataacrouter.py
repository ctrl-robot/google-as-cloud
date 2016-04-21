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
  sname='AC Router System - Kazegijutsu'
  wname='Realtime'
  dname=time.strftime('data_%Y%m%d')
  gname='Graphs'
  category = ''
  email = 'yourgooglemailaddress@gmail.com'
  password = 'yourgooglemailpassword'
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

try:
  data_wksht_id = sheets[datasheet_name]
except: #create new sheet if the required sheet doesn't exist
  client.AddWorksheet(datasheet_name,1,41,curr_key)
  feed = client.GetWorksheetsFeed(curr_key)
  sheets = get_items(feed)
  data_wksht_id = sheets[datasheet_name]
  #write header for the new sheet
  entry = client.UpdateCell(1, 1, 'machine',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 2, 'time',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 3, 'datetime',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 4, 'm01vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 5, 'm01cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 6, 'm01pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 7, 'm02vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 8, 'm02cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 9, 'm02pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 10, 'm03vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 11, 'm03cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 12, 'm03pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 13, 'm04vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 14, 'm04cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 15, 'm04pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 16, 'm05vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 17, 'm05cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 18, 'm05pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 19, 'm06vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 20, 'm06cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 21, 'm06pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 22, 'm07vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 23, 'm07cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 24, 'm07pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 25, 'm08vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 26, 'm08cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 27, 'm08pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 28, 'm09vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 29, 'm09cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 30, 'm09pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 31, 'm10vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 32, 'm10cur',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 33, 'm10pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 34, 'm11vol',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 35, 'm11cur',curr_key,data_wksht_id)  
  entry = client.UpdateCell(1, 36, 'm11pwr',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 37, 'soc1',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 38, 'f11',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 39, 'f12',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 40, 'soc2',curr_key,data_wksht_id)
  entry = client.UpdateCell(1, 41, 'f21',curr_key,data_wksht_id)
  
  if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
    print "Create new sheet succeeded."
  else:
    print "Create new sheet failed."

try:
  curr_wksht_id = sheets[worksheet_name]
except: #create new sheet if the required sheet doesn't exist
  client.AddWorksheet(worksheet_name,1,41,curr_key)
  feed = client.GetWorksheetsFeed(curr_key)
  sheets = get_items(feed)
  curr_wksht_id = sheets[worksheet_name]
  #write header for the new sheet
  entry = client.UpdateCell(1, 1, 'machine',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 2, 'time',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 3, 'datetime',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 4, 'm01vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 5, 'm01cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 6, 'm01pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 7, 'm02vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 8, 'm02cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 9, 'm02pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 10, 'm03vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 11, 'm03cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 12, 'm03pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 13, 'm04vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 14, 'm04cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 15, 'm04pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 16, 'm05vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 17, 'm05cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 18, 'm05pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 19, 'm06vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 20, 'm06cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 21, 'm06pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 22, 'm07vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 23, 'm07cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 24, 'm07pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 25, 'm08vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 26, 'm08cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 27, 'm08pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 28, 'm09vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 29, 'm09cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 30, 'm09pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 31, 'm10vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 32, 'm10cur',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 33, 'm10pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 34, 'm11vol',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 35, 'm11cur',curr_key,curr_wksht_id)  
  entry = client.UpdateCell(1, 36, 'm11pwr',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 37, 'soc1',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 38, 'f11',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 39, 'f12',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 40, 'soc2',curr_key,curr_wksht_id)
  entry = client.UpdateCell(1, 41, 'f21',curr_key,curr_wksht_id)
  
  if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
    print "Create new sheet succeeded."
  else:
    print "Create new sheet failed."

try:
  graph_wksht_id = sheets[graphsheet_name]
except: #create new sheet if the required sheet doesn't exist
  client.AddWorksheet(graphsheet_name,1,41,curr_key)
  feed = client.GetWorksheetsFeed(curr_key)
  sheets = get_items(feed)
  graph_wksht_id = sheets[graphsheet_name]
  #write header for the new sheet
  entry = client.UpdateCell(1, 1, 'time1',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 2, 'time2',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 3, 'time3',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 2, 'm01vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 5, 'm01cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 6, 'm01pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 3, 'm02vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 8, 'm02cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 9, 'm02pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 4, 'm03vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 11, 'm03cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 12, 'm03pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 5, 'm04vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 14, 'm04cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 15, 'm04pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 6, 'm05vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 17, 'm05cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 18, 'm05pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 7, 'm06vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 20, 'm06cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 21, 'm06pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 8, 'm07vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 23, 'm07cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 24, 'm07pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 9, 'm08vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 26, 'm08cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 27, 'm08pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 10, 'm09vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 29, 'm09cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 30, 'm09pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 11, 'm10vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 32, 'm10cur',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 33, 'm10pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 12, 'm11vol',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 35, 'm11cur',curr_key,graph_wksht_id)  
  entry = client.UpdateCell(1, 36, 'm11pwr',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 37, 'soc1',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 38, 'f11',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 39, 'f12',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 40, 'soc2',curr_key,graph_wksht_id)
  entry = client.UpdateCell(1, 41, 'f21',curr_key,graph_wksht_id)
  
  if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
    print "Create new sheet succeeded."
  else:
    print "Create new sheet failed."


curr_wksht=Worksheet(client,curr_key,curr_wksht_id)

graphdata=120
graph_wksht=Worksheet(client,curr_key,graph_wksht_id)
graphrow=graph_wksht.get_rows()
graphexist=len(graphrow) #

if graphexist>120:
  graph_data=120
else:
  graph_data=graphexist


#for num in range(0,graphexist)
#graph_wksht.delete_row(0)


curr_wksht.gd_client=client
curr_wksht.spreadsheet_key=curr_key
curr_wksht.worksheet_key=curr_wksht_id

usbport = '/dev/ttyAMA0'
  
ser = serial.Serial(usbport, 38400, timeout=0.1,parity=serial.PARITY_NONE, rtscts=1)
s = ser.read(512)       # read up to one hundred bytes
                        # or as much is in the buffer
line = ser.readline()   # read a '\n' terminated line
                        # write a string  ser.write("hello")

data_num=0


while True:
  #send command to pc1,pc2,pc3 to get data
  #ser.write("PC1")        # CALL PC1 TO RETRIVE DATA
  s = ser.readline(512)       # read up to one hundred bytes
  print s
  if len(s)>0:
    datastr=s.split();
    if datastr[0]=="PC3":# and datastr[9]=='\n':
      gdict = {}  
      gdict['machine'] = datastr[0]
      gdict['time'] = datastr[2]
      gdict['datetime'] = datastr[1] +" " +datastr[2]
      gdict['m01vol'] = datastr[3]
      gdict['m01cur'] = datastr[4]
      gdict['m01pwr'] = datastr[5]
      gdict['m02vol'] = datastr[6]
      gdict['m02cur'] = datastr[7]
      gdict['m02pwr'] = datastr[8]
      gdict['m03vol'] = datastr[9]
      gdict['m03cur'] = datastr[10]
      gdict['m03pwr'] = datastr[11]
      gdict['m04vol'] = datastr[12]
      gdict['m04cur'] = datastr[13]
      gdict['m04pwr'] = datastr[14]
      gdict['m05vol'] = datastr[15]
      gdict['m05cur'] = datastr[16]
      gdict['m05pwr'] = datastr[17]
      gdict['m06vol'] = datastr[18]
      gdict['m06cur'] = datastr[19]
      gdict['m06pwr'] = datastr[20]
      gdict['m07vol'] = datastr[21]
      gdict['m07cur'] = datastr[22]
      gdict['m07pwr'] = datastr[23]
      gdict['m08vol'] = datastr[24]
      gdict['m08cur'] = datastr[25]
      gdict['m08pwr'] = datastr[26]
      gdict['m09vol'] = datastr[27]
      gdict['m09cur'] = datastr[28]
      gdict['m09pwr'] = datastr[29]
      gdict['m10vol'] = datastr[30]
      gdict['m10cur'] = datastr[31]
      gdict['m10pwr'] = datastr[32]
      gdict['m11vol'] = datastr[33]
      gdict['m11cur'] = datastr[34]
      gdict['m11pwr'] = datastr[35]
      gdict['soc1'] = datastr[36]
      gdict['f11'] = datastr[37]
      gdict['f12'] = datastr[38]
      gdict['soc2'] = datastr[39]
      gdict['f21'] = datastr[40]
      
      # save history data for future analysis
      # one file per day (or a continuous execution)
      # data will be saved to this history tank for every 5 min
      # (and this freqency can be changed, and the frequency is
      # altimately controlled by the host computer where data was
      # sent from.)
      
      if graph_data%5==0:
        entry = client.InsertRow(gdict, curr_key, data_wksht_id)
        if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
          data_num=data_num+1#print "Insert row to history succeeded."
        else:
          print "Insert row to history failed."

      # prepare data for graph
      # there will be 120 data sets be save for graph. (that will be
      # about two hour data). the update period is 1 min.
      
      
      if graph_data>119:
        try:
            graph_wksht.update_row(graph_data%graphdata,gdict)
        except:
            print "update graph failed."
      else:
        entry = client.InsertRow(gdict, curr_key, graph_wksht_id)
        if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            data_num=data_num+1
        else:
            print "Insert row to graph failed."
        
      # save data for realtime display
      # realtime data will updated every one minute
      
      try:
        #curr_wksht.delete_row(0)
        curr_wksht.update_row(0,gdict)
        #data_num=data_num+1
        #print "Data updated at: "+gdict['time']+" No."+str(data_num)
        #print "Update row of realtime display succeeded."
      except:
        print "Update row of realtime display failed."
      graph_data=graph_data+1
  else:
    print "waiting for data"
  time.sleep(1)
  
