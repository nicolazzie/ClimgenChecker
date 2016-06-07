## GUI that reads and checks the CLIMGEN excel.
## Original program:
## 20/ott/2015 Gabriele Marras - Release of the python that does the checking checker
## 21/ott/2015 Ezequiel Nicolazzi - Adapted the program to a GUI
## 23/ott/2015 Ezequiel Nicolazzi - Made app for windows and macOS
## 12/jan/2016 Ezequiel Nicolazzi - Update with new template (latest update)
## 11/feb/2016 Ezequiel Nicolazzi - Added check on new variable RefSeq
##
######################################################################################################
## INTERNAL WARNING (ITALIAN):
## la lista dell'alfabeto arriva a BZ. se si inseriscono altre colonne aggiungere alla lista 'alpha'
######################################################################################################

import sys,xlrd,time
from PyQt4 import QtGui, QtCore

class MainWindow(QtGui.QMainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        self.resize(750, 350)
        self.setWindowTitle('CLIMGEN excel checker v1.5 - Latest update 11/02/16')
        widget = QtGui.QWidget(self)
        grid = QtGui.QGridLayout(widget)
        grid.setVerticalSpacing(10)
        grid.setHorizontalSpacing(8)

        button_in = QtGui.QPushButton("Open CLIMGEN excel template", widget)
        self.connect(button_in , QtCore.SIGNAL('clicked()'), self.opencheckFile)

        button_clear = QtGui.QPushButton("Clear screen", widget)
        self.connect(button_clear , QtCore.SIGNAL('clicked()'), self.clearit)

        button_save = QtGui.QPushButton("Save report to file", widget)
        self.connect(button_save , QtCore.SIGNAL('clicked()'), self.saveit)

        grid.addWidget(button_in, 0, 0)
        grid.addWidget(button_clear, 0, 1)
        grid.addWidget(button_save, 0, 2)

        self.textEdit = QtGui.QTextEdit(widget)
        self.textEdit.setReadOnly(True)
        grid.addWidget(self.textEdit, 5, 0, 1, 3)

        self.setCentralWidget(widget)

    def clearit(self):
        return self.textEdit.clear()

    def saveit(self):
        savef = QtGui.QFileDialog.getSaveFileName(self, "Save output", "Save report file", self.tr("All Files(*)"))
        if savef.isEmpty() == False:
            file = open(savef, 'w')
            file.write(str(self.textEdit.toPlainText()))
            file.close()
            self.textEdit.append('\n\n <b>Report saved in '+savef+'</b>\n')
        else:self.textEdit.append('\n\n <b>Not saving for now..</b>\n')

    def opencheckFile(self):
        go=False
        self.textEdit.clear()
        import_file = QtGui.QFileDialog.getOpenFileName(self,"Check CLIMGEN excel template", "Check new file", self.tr("XLS Files (*.xls*);;All Files (*)"))
        if not import_file or import_file.isEmpty():self.textEdit.append("Aborting opening file...\n")
        else:
            self.textEdit.append("<b>######################<br> - CLIMGEN checker v.1.5 - <br>######################</b><br><br>Opening and processing: <b>%s</b><br>" % import_file)
            go=True
        if not go:return self.textEdit.append("<font color=red>I'm still waiting for a file to chew...<b>C'mon!</b></font>")
        
        MyFile=xlrd.open_workbook(import_file)

        ##DEFINITIONS
        field_number=['alt','snpnumb','age','horncm','weightkg','witherscm',\
                 'chestcm','MEmilkyield','milkfat','milkprot','herdsize']
        field_float=['Lat','Long']
        field_accuracy=['EBVchestaccuracy','EBVudderaccuracy','EBVmilkyieldaccuracy',\
                   'EBVmilkfataccuracy','EBVmilkprotaccuracy']
        field_EBV=['EBVchest','EBVudder','EBVmilkyield',\
              'EBVmilkfat','EBVmilkprot']
        field_category={}
        field_required=['breed','sampleID','country','tissue','sex','age','DiseaseResis','Disease1','Disease2',\
                      'StateDisease1','StateDisease2','origin','geogorigin','improve','herdsize','husbandry',\
                      'siring','typicalfood','demo','herdbook','breederasso']
        field_requiredshort=['sampleID','country','tissue']
        field_condreq1=['data'] #'seqdepth','chiptype','chipname','snpnumb'
        field_condreq2=['NearestLocat']#,'Lat','Long','alt']
        field_condreq3=['species']
        field_date=['EBVdate']
        field_country=['country']
        short_check=['|7| Ovis orientalis','|8| Capra aegagrus','|9| Capra ibex','|12| Ovis vignei']
        list_of_countries=["Afghanistan", "Akrotiri", "Albania", "Algeria", "American Samoa", "Andorra", "Angola", "Anguilla", "Antarctica", "Antigua and Barbuda", "Argentina", "Armenia", "Aruba", "Ashmore and Cartier Islands", "Australia", "Austria", "Azerbaijan", "Bahamas, The", "Bahrain", "Bangladesh", "Barbados", "Bassas da India", "Belarus", "Belgium", "Belize", "Benin", "Bermuda", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Bouvet Island", "Brazil", "British Indian Ocean Territory", "British Virgin Islands", "Brunei", "Bulgaria", "Burkina Faso", "Burma", "Burundi", "Cambodia", "Cameroon", "Canada", "Cape Verde", "Cayman Islands", "Central African Republic", "Chad", "Chile", "China", "Christmas Island", "Clipperton Island", "Cocos (Keeling) Islands", "Colombia", "Comoros", "Congo, Democratic Republic of the", "Congo, Republic of the", "Cook Islands", "Coral Sea Islands", "Costa Rica", "Cote d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark", "Dhekelia", "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Ethiopia", "Europa Island", "Falkland Islands (Islas Malvinas)", "Faroe Islands", "Fiji", "Finland", "France", "French Guiana", "French Polynesia", "French Southern and Antarctic Lands", "Gabon", "Gambia, The", "Gaza Strip", "Georgia", "Germany", "Ghana", "Gibraltar", "Glorioso Islands", "Greece", "Greenland", "Grenada", "Guadeloupe", "Guam", "Guatemala", "Guernsey", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Heard Island and McDonald Islands", "Holy See (Vatican City)", "Honduras", "Hong Kong", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Isle of Man", "Israel", "Italy", "Jamaica", "Jan Mayen", "Japan", "Jersey", "Jordan", "Juan de Nova Island", "Kazakhstan", "Kenya", "Kiribati", "Korea, North", "Korea, South", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Macau", "Macedonia", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Martinique", "Mauritania", "Mauritius", "Mayotte", "Mexico", "Micronesia, Federated States of", "Moldova", "Monaco", "Mongolia", "Montserrat", "Morocco", "Mozambique", "Namibia", "Nauru", "Navassa Island", "Nepal", "Netherlands", "Netherlands Antilles", "New Caledonia", "New Zealand", "Nicaragua", "Niger", "Nigeria", "Niue", "Norfolk Island", "Northern Mariana Islands", "Norway", "Oman", "Pakistan", "Palau", "Panama", "Papua New Guinea", "Paracel Islands", "Paraguay", "Peru", "Philippines", "Pitcairn Islands", "Poland", "Portugal", "Puerto Rico", "Qatar", "Reunion", "Romania", "Russia", "Rwanda", "Saint Helena", "Saint Kitts and Nevis", "Saint Lucia", "Saint Pierre and Miquelon", "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia and Montenegro", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Georgia and the South Sandwich Islands", "Spain", "Spratly Islands", "Sri Lanka", "Sudan", "Suriname", "Svalbard", "Swaziland", "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tokelau", "Tonga", "Trinidad and Tobago", "Tromelin Island", "Tunisia", "Turkey", "Turkmenistan", "Turks and Caicos Islands", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States", "Uruguay", "Uzbekistan", "Vanuatu", "Venezuela", "Vietnam", "Virgin Islands", "Wake Island", "Wallis and Futuna", "West Bank", "Western Sahara", "Yemen", "Zambia", "Zimbabwe"]
        allowed_countries={x.lower():'0' for x in list_of_countries}

        alpha=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X',\
              'Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ',\
              'AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ',\
              'BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ']

        ##SHEET VARs
        observation= MyFile.sheet_by_index(1)
        category= MyFile.sheet_by_index(0)

        for rx in range(category.ncols):
            list=[x for x in category.col_values(rx) if x != ''] #delete value empty''
            field_category[category.col(rx)[0].value]=list       #dictionary

        ### Check file
        self.textEdit.append('<br><b>Reporting errors found...</b>')
        self.textEdit.append('<font color=black>SAMPLE_ID;ROW;COL;HEADER;VALUE;TYPE_OF_ERROR</font>')

        animal=[];err=False
        for line in range(observation.nrows):
            #Creating list with header
            if line==0: 
                header=[observation.row(line)[x].value for x in range(len(observation.row(line)))]
                continue
            riga=(str(line+1)) #File line (row)
            sample=(observation.row(line)[2].value) #sampleID
            ### CHECK for double samples
            if sample in animal and sample != '': ##skip empty
                err=True
                self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;SAMPLE ID REPEATED</font>' \
                            % (sample,riga,alpha[2],header[2],sample))
            animal.append(sample)

            #Start checking by line
            for val in range(len(observation.row(line))):
                ## EXCEPTION 1 (Field "data", conditional to what kind of data there is)
                if header[val] in field_condreq1:
                    #DATA is a required field. So if it is missing, this is an error.   
                    if observation.row(line)[val].value=='':
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a1)</font>' \
                                       % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                        continue
                    #DATA is also a dictionary with catheogrical variables. So, if the variable does not belong, it is an error                      
                    if field_category.has_key(header[val]): #list value dictionary
                        if not observation.row(line)[val].value in field_category[header[val]]:
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;VALUE NOT ALLOWED ON THIS VARIABLE (Err_b1)</font>' \
                                                % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                            continue
                    #If data = SNPchip, then there's no meaning on having seqdepth as a reqd variable
                    if observation.row(line)[val].value=='|4| SNP-Chip' and (observation.row(line)[val+1].value=='' or observation.row(line)[val+2].value=='' or observation.row(line)[val+3].value==''):
                        err=True
                        for z in range(1,4,1):
                            if observation.row(line)[val+z].value=='':
                                self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a2)</font>' \
                                       % (sample,riga,alpha[val+z],header[val+z],observation.row(line)[val+z].value))
                        continue
                    if observation.row(line)[val].value=='|2| GBS' and observation.row(line)[val+3].value=='':
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a3)</font>' \
                               % (sample,riga,alpha[val+3],header[val+3],observation.row(line)[val+3].value))
                        continue
                    if observation.row(line)[val].value!='|4| SNP-Chip':
                        if observation.row(line)[val+4].value=='':
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a4)</font>' \
                                       % (sample,riga,alpha[val+4],header[val+4],observation.row(line)[val+4].value))
            
                        if observation.row(line)[val+5].value=='':
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a5) </font>' \
                                       % (sample,riga,alpha[val+5],header[val+5],observation.row(line)[val+5].value))
                        continue

                ## EXCEPTION 2
                elif header[val] in field_condreq2:
                    #One is a reqd field: not all can be missing   
                    #Lat,Long,alt,NearestLocat,largecity  #NearestLocat AND largecity   must be present if no GIS is provided.
                    if observation.row(line)[val].value=='' and observation.row(line)[val+1].value=='':
                        if observation.row(line)[val-3].value=='' and observation.row(line)[val-2].value=='' and observation.row(line)[val-1].value=='':
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (NearestLocat+largecity OR full GIS coords are reqd) (Err_a6)</font>' \
                                   % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                            continue
                        if observation.row(line)[val-3].value=='' or observation.row(line)[val-2].value=='' or observation.row(line)[val-1].value=='':
                            err=True
                            for z in range(1,4,1):
                                if observation.row(line)[val-z].value=='':
                                    self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (NearestLocat+largecity OR full GIS coords are reqd) (Err_a7)</font>' \
                                            % (sample,riga,alpha[val-z],header[val-z],observation.row(line)[val-z].value))
                        continue
                    else:
                        for z in range(1,2,1):
                            if observation.row(line)[val+z].value=='':
                                self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (NearestLocat+largecity OR full GIS coords are reqd) (Err_a8)</font>' \
                                    % (sample,riga,alpha[val-z],header[val-z],observation.row(line)[val-z].value))
                        continue
                        
                ### REQUIRED FIELDS (SHORT)
                if observation.row(line)[0].value in short_check:
                    if header[val] in field_condreq3: 
                        if observation.row(line)[val].value=='':
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a9)</font>' \
                                       % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                            continue
                        if field_category.has_key(header[val]): #list value dictionary
                            if not observation.row(line)[val].value in field_category[header[val]]:
                                err=True
                                self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;VALUE NOT ALLOWED ON THIS VARIABLE (Err_b2)</font>' \
                                       % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                    
                else:    
                ### REQUIRED FIELDS (FULL)
                    if header[val] in field_required:
                        if observation.row(line)[val].value=='':
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;REQUIRED FIELD IS MISSING (Err_a10)</font>' \
                                       % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                            continue
                        if field_category.has_key(header[val]): #list value dictionary
                            if not observation.row(line)[val].value in field_category[header[val]]:
                                err=True
                                self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;VALUE NOT ALLOWED ON THIS VARIABLE (Err_b3)</font>' \
                                       % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                ### CATEGORIES CHECK - ALWAYS
                if header[val] in field_category:
                    if observation.row(line)[val].value=='':continue #skip se la variabile e vuota
                    if field_category.has_key(header[val]): #list value dictionary
                        if not observation.row(line)[val].value in field_category[header[val]]:
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;VALUE NOT ALLOWED ON THIS VARIABLE (Err_b4)</font>' \
                                                % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))

                if observation.row(line)[val].value=='':continue #skip se la variabile e vuota        
                elif header[val] in field_number: ##number
                    if not type(observation.row(line)[val].value) == float:
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;IS NOT A NUMBER (Err_c1)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                    elif observation.row(line)[val].value <= 0:
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;OUT OF RANGE (n> 0)  (Err_d1)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))

                elif header[val] in field_float: ##float number
                    if not type(observation.row(line)[val].value) == float:
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;IS NOT A NUMBER (Err_c2)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                    elif -180 > observation.row(line)[val].value or observation.row(line)[val].value > 180: 
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;OUT OF RANGE (-180 <n< 180)  (Err_d2)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))

                elif header[val] in field_accuracy: ##EBVaccuracy
                    if not type(observation.row(line)[val].value) == float:
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;IS NOT A NUMBER (Err_c3)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                    elif 0 > observation.row(line)[val].value or observation.row(line)[val].value > 1: 
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;OUT OF RANGE (0 <n< 1) (Err_d3)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))

                elif header[val] in field_EBV: ##EBV number
                    if not type(observation.row(line)[val].value) == float:
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;IS NOT A NUMBER (Err_c4)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                    elif -99999 > observation.row(line)[val].value or observation.row(line)[val].value > 99999: 
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;OUT OF RANGE (-99999 <n< 99999)  (Err_d4)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))

                elif header[val] in field_date: ##date
                    data_cell=observation.row(line)[val].value
                    if not '/' in data_cell or len(data_cell) != 7:
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;DATA WRONG FORMAT (SHOULD BE MM/YYYY) (Err_e1)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                    else:
                        if int(data_cell[0:2]) in range(1,12) or int(data_cell[3:7]) in range(1000,3000):
                            err=True
                            self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s;DATA OUT OF RANGE (1 <MM< 12 and 1000 <YYYY< 3000) (Err_d5)</font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))

                elif header[val] in field_country: ##country 
                    data_cell=observation.row(line)[val].value.lower()
                    if not allowed_countries.has_key(data_cell):
                        err=True
                        self.textEdit.append('<font color=orange>%s;%s;%s;%s;%s; UNRECOGNISED COUNTRY (see manual) (Err_f1)<\font>' \
                                    % (sample,riga,alpha[val],header[val],observation.row(line)[val].value))
                        continue
        if line==0:
            err=True
            self.textEdit.append('<font color=orange>Template file is empty! Please check and run again</font>')
        if err:
            self.textEdit.append('<br><font color=red>#######################################################</font>')
            self.textEdit.append('<font color=red><b>      ERRORS FOUND IN FILE - PLEASE CORRECT AND RUN AGAIN    </b></font>')
            self.textEdit.append('<font color=red>#######################################################</font>')
        else:
            self.textEdit.append('<br><font color=green>NO ERRORS FOUND IN<b> %s </b>LINES READ</font><br>' % str(line))
            self.textEdit.append('<font color=green>######################<br># <b>CONGRATULATIONS! </b>#<br>######################</font><br>')
            self.textEdit.append('Template: <b>%s<b> <font color=green>CHECKED OK!</font><br>' % import_file) 
            self.textEdit.append('<b>You can now transfer this excel file to the CLIMGEN server</b>')
                                    

app = QtGui.QApplication(sys.argv)
main = MainWindow()
main.show()
main.raise_()
sys.exit(app.exec_())

