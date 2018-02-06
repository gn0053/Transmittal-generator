''' 
VERSION 4.0
Requirements to run this:
    python 2.7
    PyQt Library
    Jinja library
    docxtpl library
    Use pip to intall libraries
    
Note: the '.self' you see per functions(def) defines that the function under that class belongs to it (parent class)
        the same can be said with the 'self.variable', and thus can be accesed to any part of the code under that class
        take note not to just declare everything in 'self.variable' to prevent code spagettification
        
        Due to the nature of Python, you can add modules or just comment modules to disable them,
        you can also easily add functions and classes without breaking the application

'''
#core imports
#builtin in python
ver = '4.1'
import sys
import math
import os.path
import os, datetime
import ConfigParser
import re
import string
import shutil
#end of builtin imports

#from pyqt library
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from PyQt4.QtWebKit import *
from PyQt4 import QtGui, QtCore
#end of pyqt library


from docxtpl import * #uses .docx templates instead of .html
from jinja2 import Environment, FileSystemLoader # for ease of doc generation (refer to django templating)

#subclassed classes


#human sort
convert = lambda text: int(text) if text.isdigit() else text
alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key) ]
#end of human sort

    
#subcalssed QDialog, enables the window to catch events
class option_win(QDialog):
    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, 'Close',
            "Are you Sure? \nAll your Inputs will be lost.", QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

class def_list_dialog(QDialog):
    keyPressed = QtCore.pyqtSignal()

    def keyPressEvent(self, event):
        if (event.type()==QEvent.KeyPress) and (event.key()==Qt.Key_Delete):
            self.keyPressed.emit()
        super(QDialog, self).keyPressEvent(event)
        
class Window(QtGui.QMainWindow):
    
    def __init__(self, parent=None):         #initialization of the code under it
        super(Window, self).__init__(parent) 
        self.setGeometry(750, 450, 400, 200)
        self.setWindowTitle("Transmital Generator Build ver. " + ver)
        self.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        self.setFixedSize(self.size())
        
        #dropdown options
        extractAction1 = QtGui.QAction("&Documentation", self)
        extractAction1.triggered.connect(self.open_about)
        
        extractAction2 = QtGui.QAction("&Usage", self)
        extractAction2.triggered.connect(self.open_docs)
        
        extractAction3 = QtGui.QAction("&Change Log", self)
        extractAction3.triggered.connect(self.open_ver)
        
        extractAction4 = QtGui.QAction("Add Job Defaults", self)
        extractAction4.setToolTip("Add/Set Job Defaults")
        extractAction4.triggered.connect(self.list_of_defaults)
        #end of options
        
        mainMenu = self.menuBar()
        fileMenu1 = mainMenu.addMenu('&Help')
        fileMenu1.addAction(extractAction1)
        fileMenu1.addAction(extractAction2)
        fileMenu1.addAction(extractAction3)
        
        fileMenu2 = mainMenu.addMenu('&Options')
        fileMenu2.addAction(extractAction4)
        
        
        self.home()
    #end of main window
       
    #main window extension
    
    def home(self):
        job_name = self.get_def_list()
        
        groupBox = QGroupBox("Job Selection", self)
        groupBox.setGeometry(QRect(10, 70, 380, 100))
        
        self.combo = QtGui.QComboBox(self)
        self.combo.setToolTip("Selection of Jobs")
        job_name.sort(reverse = True)
        self.combo.addItems(job_name)
        
        vbox = QGridLayout()
        vbox.addWidget(self.combo, 0,0) 
        groupBox.setLayout(vbox)
        
        self.toolBar = self.addToolBar("Strict Filter")
        
        self.filter_field = QLineEdit(self)
        self.filter_field.setFrame(False)
        
        self.icon = QLabel(self)
        self.icon.setPixmap(QPixmap("media/if_Zoom.png"))
        self.icon.setStyleSheet("""
        .QWidget {
            background-color: white;
            }
        """)
        
        self.filter_field.setPlaceholderText("Filter")
        self.filter_field.setToolTip("Will Search based on matches")
        self.toolBar.addWidget(self.icon)
        self.toolBar.addWidget(self.filter_field)
        
        btn1 = QtGui.QPushButton("OK", self)
        btn1.clicked.connect(lambda: self.trans_gen_dir(unicode(self.combo.currentText())))
        btn1.resize(btn1.minimumSizeHint())
        btn1.move(205,170)
        btn1.setToolTip("Select Current Job Displayed")
        btn1.setFocusPolicy(Qt.StrongFocus)
        btn1.setDefault(True)
        btn1.setAutoDefault(True)
        
        self.filter_field.textChanged.connect(lambda: self.action_change(job_name))
        
        btn2 = QtGui.QPushButton("Exit", self)
        btn2.clicked.connect(self.close_application)
        btn2.resize(btn2.minimumSizeHint())
        btn2.move(299,170)
        btn2.setToolTip("Exit Application")
        btn2.setAutoDefault(True)
        
        btn3 = QtGui.QPushButton("Refresh", self)
        btn3.clicked.connect(self.refresh_list)
        btn3.resize(btn2.minimumSizeHint())
        btn3.move(10,170)
        self.show()
        
    #end of extension
    
    def cell_fetch(self, item):
        g = ''
        if not item:
            row = self.table.currentRow()
            g = self.table.item(row, 0).text()
        else:
            g = item.text()
        self.set_defaults(unicode(g))
        
    def del_default(self):
        row = self.table.currentRow()
        g = self.table.item(row, 0).text()
        reply = QtGui.QMessageBox.question(self.table, str(g),
            "Are you Sure? \nAll your defaults will be lost.", QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            shutil.rmtree(os.path.join('job_defaults', str(g)))
            self.populate_table()
        
    def list_of_defaults(self):
        self.def_list = def_list_dialog()
        self.def_list.resize(800, 865)
        self.def_list.setFixedSize(self.def_list.size())
        self.def_list.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.def_list.setWindowTitle("Job Defaults")
        self.def_list.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        self.table = QTableWidget()
        self.table.resize(400, 300)
        self.def_list.installEventFilter(self)
        groupBox1 = QGroupBox("", self.def_list)
        groupBox1.setGeometry(QRect(10, 10, 780, 710))
        vbox = QGridLayout()
        
        vbox.addWidget(self.table, 0,0)
        groupBox1.setLayout(vbox)
        
        btn2 = QtGui.QPushButton("Close", self.def_list)
        btn2.clicked.connect(self.def_list_close)
        btn2.resize(btn2.minimumSizeHint())
        btn2.move(700,830)
        
        btn3 = QtGui.QPushButton("Edit", self.def_list)
        btn3.clicked.connect(self.cell_fetch)
        btn3.resize(btn2.minimumSizeHint())
        btn3.move(600,830)
        
        btn4 = QtGui.QPushButton("Add", self.def_list)
        btn4.clicked.connect(self.get_job)
        btn4.resize(btn2.minimumSizeHint())
        btn4.move(500,830)
        
        btn5 = QtGui.QPushButton("Delete", self.def_list)
        btn5.clicked.connect(self.del_default)
        btn5.resize(btn2.minimumSizeHint())
        btn5.move(10,830)
        
        self.table.setColumnCount(1)
        self.table.setHorizontalHeaderLabels(('Jobs',))
        self.table.itemDoubleClicked.connect(self.cell_fetch)
        self.table.clearSelection()
        self.def_list.keyPressed.connect(self.del_default)
        self.populate_table()
        self.def_list.show()
        
    def get_def_list(self):
        job_name = []
        for f in os.listdir('job_defaults'):
            if os.path.isdir(os.path.join('job_defaults', f)):
                job_name += [f]  
        return job_name
        
    def refresh_list(self):
        job_list = self.get_def_list()
        self.combo.clear()
        job_list.sort(reverse = True)
        self.combo.addItems(job_list)
        
    def populate_table(self):
        job_name = self.get_def_list()
        index_counter = 0
        self.table.setRowCount(len(job_name))
        for items in job_name:
            cell_item = QtGui.QTableWidgetItem(items)
            cell_item.setFlags(Qt.ItemIsSelectable|Qt.ItemIsEnabled)
            self.table.setItem(index_counter,0, cell_item)
            index_counter += 1
        header = self.table.horizontalHeader()
        header.setResizeMode(QHeaderView.Fixed)
        self.table.setColumnWidth(0, 736)
        self.def_list.setWindowModality(QtCore.Qt.ApplicationModal)
        
        
    def get_job(self):
        job = QtGui.QFileDialog.getExistingDirectory(self.def_list, 'Select Job Path')
        
        if job:
            self.new_def(unicode(job.split('\\')[len(job.split('\\'))-1]), unicode(job))
        else:
            self.new_def('','')
            
    def new_def(self, job, job_dir):
        if len(job) > 0 and len(job_dir) > 0:
            if not os.path.exists(os.path.join('job_defaults', job)):
                proj_num = ''
                trans_num = ''
                trans_list = []
                trans_folder = ''
                special_char = '!@#$%^&*()_+=-{}[]|"\'?/><, '
                trans_num2 = ''
                for folders in os.listdir(job_dir):
                    if 'TRANSMITTAL' in folders or 'Transmittal' in folders or 'transmittal' in folders:
                        trans_folder = folders
                trans_list = os.listdir(os.path.join(job_dir, trans_folder))
                if trans_folder:
                    for files in range(len(trans_list)):
                        trans_num1 = ''
                        trans_on = False
                        get_job_num = False
                        for char in trans_list[files]:
                            if char == '_':
                                get_job_num = False
                                
                            if get_job_num == True:
                                if char.isdigit():
                                    job_num += char
                            if unicode(char) == '#':
                                trans_on = True
                                trans_num1 = ''
                                
                            if (not char.isdigit() and char != '#' and trans_num1) or (len(trans_list) == files):
                                trans_on = False
                                if int(trans_num1) >= int(trans_num2):
                                    trans_num2 = trans_num1
                                    trans_num1 = ''
                            if trans_on == True and char.isdigit():
                                trans_num1 += char
                                if not trans_num2:
                                    trans_num2 += char
                    if trans_num2 and len(trans_num2) > 0:           
                        trans_num = trans_num2
                    else:
                       trans_num = '000' 
                        
                else:
                    trans_num = '000'
                    
                for char in job:
                    if char in special_char:
                        break
                    else:
                        proj_num += char
                        
                os.makedirs(os.path.join('job_defaults', job))
                f = open(os.path.join('job_defaults', job, job + ".ini"),"w+")
                f.write('[defaults]\n')
                f.write('to =\n')
                f.write('Transmittal# = ' + trans_num + '\n')
                f.write('Project# = ' + proj_num + '\n')
                f.write('project =\n')
                f.write('location =\n')
                f.write('proj_man =\n')
                f.write('attn =\n')
                f.close()
                self.populate_table()
                self.set_defaults(job)
                
            else:
                reply = QtGui.QMessageBox.question(self, 'Add New',
                "Job Already Have a Default, Edit Instead?", QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                
                if reply == QtGui.QMessageBox.Yes:
                    self.set_defaults(job)
                    
        
    def set_defaults(self, job):# The window for editing default.ini
        
        config = ConfigParser.ConfigParser()
        config.read(os.path.join('job_defaults', job, job + '.ini'))
        trans_num = config.get('defaults','Transmittal#')
        to = config.get('defaults','to')
        proj = config.get('defaults','Project')
        proj_num = config.get('defaults','Project#')
        loc = config.get('defaults','location')
        proj_man = config.get('defaults','proj_man')
        attn = config.get('defaults','attn')
        
        
        self.default_win = QDialog()
        self.default_win.resize(400, 450)
        self.default_win.setFixedSize(self.default_win.size())
        self.default_win.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.default_win.setWindowTitle(job)
        self.default_win.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        self.default_win.setWindowModality(QtCore.Qt.ApplicationModal)
        
        proj_num_lbl = proj_label = QLabel(self.default_win)
        proj_num_lbl.setText("Project Number:")
        proj_num_lbl.move(10,15)
        
        proj_num_def = QtGui.QLineEdit(self.default_win)
        proj_num_def.move(160,10)
        proj_num_def.resize(100, 30)
        proj_num_def.setText(proj_num)
        
        trans_num_lbl = proj_label = QLabel(self.default_win)
        trans_num_lbl.setText("Next Transmittal Number:")
        trans_num_lbl.move(10,55)
        
        trans_num_def = QtGui.QLineEdit(self.default_win)
        trans_num_def.move(160,50)
        trans_num_def.resize(100, 30)
        trans_num_def.setText(trans_num)
        
        fname_label = QLabel(self.default_win)
        fname_label.setText("Job Name:")
        fname_label.move(10,105)
        
        fname_def = QtGui.QLineEdit(self.default_win)
        fname_def.move(75,100)
        fname_def.resize(295, 30)
        fname_def.setText(job)
        
        attn_label = QLabel(self.default_win)
        attn_label.setText("Attn:")
        attn_label.move(10,160)
        
        attn_def = QtGui.QLineEdit(self.default_win)
        attn_def.move(70,155)
        attn_def.resize(300, 30)
        attn_def.setText(attn)
        
        proj_label = QLabel(self.default_win)
        proj_label.setText("Project:")
        proj_label.move(10,195)
        
        proj_def = QtGui.QLineEdit(self.default_win)
        proj_def.move(70,190)
        proj_def.resize(300, 30)
        proj_def.setText(proj)
        
        loc_label = QLabel(self.default_win)
        loc_label.setText("Location:")
        loc_label.move(10,230)
        
        def_loc = QtGui.QLineEdit(self.default_win)
        def_loc.move(70,225) 
        def_loc.resize(300, 30)
        def_loc.setText(loc)
        
        proj_man_label = QLabel(self.default_win)
        proj_man_label.setText("Project\nManager:")
        proj_man_label.move(10,260)
        
        def_proj_man = QtGui.QLineEdit(self.default_win)
        def_proj_man.move(70,260)
        def_proj_man.resize(300, 30)
        def_proj_man.setText(proj_man)
        
        to_whom_label = QLabel(self.default_win)
        to_whom_label.setText("TO:")
        to_whom_label.move(10,300)
        
        def_to_whom = QtGui.QTextEdit(self.default_win)
        def_to_whom.resize(300, 100)
        def_to_whom.move(70,300)
        def_to_whom.setPlainText(to)
        
        def_save_btn = QtGui.QPushButton("Save", self.default_win)
        def_save_btn.clicked.connect(lambda: 
                                    self.save_def(job, proj_num_def, 
                                    trans_num_def, proj_def, def_loc, 
                                    def_proj_man, def_to_whom, fname_def, attn_def))
                                    
        def_save_btn.resize(def_save_btn.minimumSizeHint())
        def_save_btn.move(210,420)
        
        cancel_btn = QtGui.QPushButton("Exit", self.default_win)
        cancel_btn.clicked.connect(self.def_close)
        cancel_btn.resize(cancel_btn.minimumSizeHint())
        cancel_btn.move(305,420)
        
        self.default_win.show()
        
    def save_def(self, job, proj_num_def, trans_num_def, proj_def, def_loc, def_proj_man, def_to_whom, fname_def, attn_def):
        proj_num_def = str(proj_num_def.text())
        trans_num_def = str(trans_num_def.text())
        proj_def = str(proj_def.text())
        def_loc = str(def_loc.text())
        def_proj_man = str(def_proj_man.text())
        def_to_whom = str(def_to_whom.toPlainText())
        def_attn = str(attn_def.text())
        save_suc = False
        if os.path.isfile(os.path.join('job_defaults', str(job), str(job) + '.ini')):
            def_ini = ConfigParser.ConfigParser()
            def_ini.read(os.path.join('job_defaults', str(job), str(job) + '.ini'))
            if not def_ini.has_section('defaults'):
                def_ini.add_section('defaults')
                def_ini.set('defaults', 'Project#', proj_num_def)
                def_ini.set('defaults', 'Transmittal#', trans_num_def)
                def_ini.set('defaults', 'to', def_to_whom)
                def_ini.set('defaults', 'project', proj_def)
                def_ini.set('defaults', 'location', def_loc)
                def_ini.set('defaults', 'proj_man', def_proj_man)
                def_ini.set('defaults', 'attn', def_attn)
                with open(os.path.join('job_defaults', job, job + '.ini'), 'wb') as configfile:
                    def_ini.write(configfile)
                self.infmsg("Saved", "Defaults saved")
                
            else:
                def_ini.set('defaults', 'Project#', proj_num_def)
                def_ini.set('defaults', 'Transmittal#', trans_num_def)
                def_ini.set('defaults', 'to', def_to_whom)
                def_ini.set('defaults', 'project', proj_def)
                def_ini.set('defaults', 'location', def_loc)
                def_ini.set('defaults', 'proj_man', def_proj_man)
                def_ini.set('defaults', 'attn', def_attn)
                with open(os.path.join('job_defaults', job, job + '.ini'), 'wb') as configfile:
                    def_ini.write(configfile)
                save_suc = True
        else:
            f = open(os.path.join('job_defaults', job, job + ".ini"),"w+")
            f.write('[defaults]\n')
            f.write('to = ' + def_to_whom + '\n')
            f.write('Transmittal# = ' + trans_num_def + '\n')
            f.write('Project# = ' + proj_num_def + '\n')
            f.write('project = ' + proj_def + '\n')
            f.write('location = ' + def_loc + '\n')
            f.write('proj_man = ' + def_proj_man + '\n')
            f.write('attn = ' + def_proj_man + '\n')
            f.close()
            
        if unicode(fname_def.text()) != job:
            reserved_char = '<>:"/\\|?*.'
            err = False
            invalid_char = ''
            for char in  unicode(fname_def.text()):
                if char in reserved_char:
                    err = True
                    invalid_char += char
                    
            if not err and save_suc:
                old_path = os.path.join('job_defaults', str(job))
                new_path = os.path.join('job_defaults', str(fname_def.text()))
                
                os.rename(os.path.join(old_path, job + '.ini'), os.path.join(old_path, str(fname_def.text()) + '.ini'))
                os.rename(old_path, new_path)
                self.populate_table()
                self.combo.clear()
                self.combo.addItems(self.get_def_list())
                job = unicode(fname_def.text())
                self.set_defaults(job)
                self.infmsg("Saved", "Defaults saved")
            else:
                self.errmsg("Invalid Character detected: " + invalid_char)
        else:
            if save_suc:
                self.infmsg("Saved", "Defaults saved")
                
    def trans_gen_dir(self, job):
        if job and len(job) > 0:
            file = QtGui.QFileDialog.getExistingDirectory(self, 'Select Job Path')
            if file:
                if os.path.exists(file):
                    self.def_fetch(file, job)
                else:
                    self.errmsg("Path Does Not Exist")
            else:
                self.errmsg("No Path Selected")
        else:
            self.warnmsg("No job selected")
            
    def def_fetch(self, dir, job):
        config = ConfigParser.ConfigParser()
        config.read(os.path.join('job_defaults', job, job + '.ini'))
        trans_num_def = config.get('defaults','Transmittal#')
        def_to_whom = config.get('defaults','to')
        proj_def = config.get('defaults','Project')
        proj_num_def = config.get('defaults','Project#')
        def_loc = config.get('defaults','location')
        def_proj_man = config.get('defaults','proj_man')
        def_attn = config.get('defaults','attn')
        self.Options_box(dir, proj_num_def, trans_num_def, proj_def, def_loc, def_proj_man, def_to_whom, job, def_attn)
        
    def Options_box(self, dir, proj_num_def, trans_num_def, proj_def, def_loc, def_proj_man, def_to_whom, job, def_attn): #user input for the transmittal file to be generated
        
        td = datetime.datetime.now().strftime('%b %d, %Y')
        self.option_win = option_win(self)
        self.option_win.resize(800, 865)
        self.option_win.setFixedSize(self.option_win.size())
        self.option_win.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.option_win.setWindowTitle("Output settings")
        self.option_win.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        self.option_win.setWindowModality(QtCore.Qt.ApplicationModal)
        
        groupBox1 = QGroupBox("Basic Information", self.option_win)
        groupBox1.setGeometry(QRect(10, 10, 780, 155))
        self.groupBox1 = groupBox1
        
        label1 = QLabel()
        label1.setText("Transmittal #")
        self.label1 = label1
        
        trans = QLineEdit()
        trans.setText(trans_num_def)
        self.trans = trans
        
        label2 = QLabel()
        label2.setText("Date")
        self.label2 = label2
        
        dte = QLineEdit(self.option_win)
        dte.setText(unicode(td))
        self.dte = dte
        
        label3 = QLabel()
        label3.setText("Project #")
        self.label3 = label3
        proj_num = QLineEdit(self.option_win)
        proj_num.setText(proj_num_def)
        self.proj_num = proj_num
        
        label4 = QLabel()
        self.label4 = label4
        label4.setText("Project")
        proj = QtGui.QLineEdit(self.option_win)
        proj.setText(proj_def)
        self.proj = proj
        
        label5 = QLabel()
        label5.setText("Location")
        self.label5 = label5
        loc = QtGui.QLineEdit(self.option_win)
        loc.setText(def_loc)
        self.loc = loc
        
        label6 = QLabel()
        label6.setText("TO: ")
        self.label6 = label6
        to_whom = QtGui.QTextEdit(self.option_win)
        to_whom.setText(def_to_whom)
        self.to_whom = to_whom
        
        vbox = QGridLayout()
        
        vbox.addWidget(label1, 0,0)
        vbox.addWidget(trans, 0,1)
        vbox.addWidget(label2, 0,2)
        vbox.addWidget(dte, 0,3)
        vbox.addWidget(label3, 0,4)
        vbox.addWidget(proj_num, 0,5)
        
        vbox.addWidget(label4, 1,0)
        vbox.addWidget(proj, 1,1)
        vbox.addWidget(label5, 1,2)
        vbox.addWidget(loc, 1,3)
        vbox.addWidget(label6, 2,0)
        vbox.addWidget(to_whom, 2,1)
        
        groupBox1.setLayout(vbox)
        
        groupBox2 = QGroupBox("With/Separate:", self.option_win)
        groupBox2.setGeometry(QRect(10, 170, 780, 100))
        vbox2 = QGridLayout()
        op1 = QCheckBox("Herewith")
        op2 = QCheckBox("Under separate cover via: ")
        desc = QtGui.QLineEdit(self.option_win)
        self.desc = desc
        self.op1 = op1
        self.op2 = op2
        self.desc = desc
        self.groupBox2 = groupBox2
        
        attnlabel = QLabel()
        attnlabel.setText("Attn:")
        attn = QtGui.QLineEdit(self.option_win)
        attn.setText(def_attn)
        self.attn = attn
        vbox2.addWidget(op1, 0,0)
        vbox2.addWidget(op2, 0,1)
        vbox2.addWidget(desc, 0,2)
        vbox2.addWidget(attnlabel, 0,3)
        vbox2.addWidget(attn, 0,4)
        groupBox2.setLayout(vbox2)
        
        
        groupBox3 = QGroupBox("Files To Be Sent", self.option_win)
        groupBox3.setGeometry(QRect(10, 280, 780, 130))
        self.groupBox3 = groupBox3
        vbox3 = QGridLayout()
        
        file_op1 = QtGui.QCheckBox('Shop Drawings', groupBox3)
        file_op2 = QtGui.QCheckBox('Prints', groupBox3)
        file_op3 = QtGui.QCheckBox('Data Sheets', groupBox3)
        
        file_op4 = QtGui.QCheckBox('Copy of Letter', groupBox3)
        file_op5 = QtGui.QCheckBox('Change Order', groupBox3)
        file_op6 = QtGui.QCheckBox('Schedule', groupBox3)
        
        file_op7 = QtGui.QCheckBox('Drawings', groupBox3)
        file_op8 = QtGui.QCheckBox('Specifications', groupBox3)
        file_op9 = QtGui.QCheckBox('Addendum #', groupBox3)
        
        file_op10 = QtGui.QCheckBox('Samples', groupBox3)
        file_op11 = QtGui.QCheckBox('Brochure', groupBox3)
        file_op12 = QtGui.QCheckBox('Catalog Cuts', groupBox3)
        
        self.file_op1 = file_op1
        self.file_op2 = file_op2
        self.file_op3 = file_op3
        self.file_op4 = file_op4
        self.file_op5 = file_op5
        self.file_op6 = file_op6
        self.file_op7 = file_op7
        self.file_op8 = file_op8
        self.file_op9 = file_op9
        self.file_op10 = file_op10
        self.file_op11 = file_op11
        self.file_op12 = file_op12
        
        vbox3.addWidget(file_op1, 0,0)
        vbox3.addWidget(file_op2, 1,0)
        vbox3.addWidget(file_op3, 2,0)
        
        vbox3.addWidget(file_op4, 0,1)
        vbox3.addWidget(file_op5, 1,1)
        vbox3.addWidget(file_op6, 2,1)
        
        vbox3.addWidget(file_op7, 0,2)
        vbox3.addWidget(file_op8, 1,2)
        vbox3.addWidget(file_op9, 2,2)
        
        vbox3.addWidget(file_op10, 0,3)
        vbox3.addWidget(file_op11, 1,3)
        vbox3.addWidget(file_op12, 2,3)
        
        groupBox3.setLayout(vbox3)
        
        groupBox5 = QGroupBox("Purpose", self.option_win)
        groupBox5.setGeometry(QRect(10, 420, 780, 180))
        self.groupBox5 = groupBox5
        vbox5 = QGridLayout()
        
        stat_op1 = QtGui.QCheckBox('For Approval', groupBox5)
        stat_op2 = QtGui.QCheckBox('For Re-Approval', groupBox5)
        stat_op3 = QtGui.QCheckBox('Revised and Re-Submit', groupBox5)
        stat_op4 = QtGui.QCheckBox('For Fabrication', groupBox5)
        stat_op5 = QtGui.QCheckBox('Revised for Fabrication', groupBox5)
        stat_op6 = QtGui.QCheckBox('For construction', groupBox5)
        
        stat_op7 = QtGui.QCheckBox('Revised for Construction', groupBox5)
        stat_op8 = QtGui.QCheckBox('For Approval/Fabrication', groupBox5)
        stat_op9 = QtGui.QCheckBox('For Field Works', groupBox5)
        stat_op10 = QtGui.QCheckBox('For Review & Comment', groupBox5)
        stat_op11 = QtGui.QCheckBox('Record Copy', groupBox5)
        stat_op12 = QtGui.QCheckBox('As Requested', groupBox5)
        
        stat_op13 = QtGui.QCheckBox('Specifications', groupBox5)
        stat_op14 = QtGui.QCheckBox('Addendum #', groupBox5)
        stat_op15 = QtGui.QCheckBox('For Quotation Due', groupBox5)
        stat_op16 = QtGui.QCheckBox('Reviewed by CORE', groupBox5)
        stat_op17 = QtGui.QCheckBox('Reviewed by QA/QR', groupBox5)
        
        vbox5.addWidget(stat_op1, 0,0)
        vbox5.addWidget(stat_op2, 1,0)
        vbox5.addWidget(stat_op3, 2,0)
        vbox5.addWidget(stat_op4, 3,0)
        vbox5.addWidget(stat_op5, 4,0)
        vbox5.addWidget(stat_op6, 5,0)
        
        vbox5.addWidget(stat_op7, 0,1)
        vbox5.addWidget(stat_op8, 1,1)
        vbox5.addWidget(stat_op9, 2,1) 
        vbox5.addWidget(stat_op10, 3,1)
        vbox5.addWidget(stat_op11, 4,1)
        vbox5.addWidget(stat_op12, 5,1)
        
        vbox5.addWidget(stat_op13, 0,2)
        vbox5.addWidget(stat_op14, 1,2)
        vbox5.addWidget(stat_op15, 2,2)
        vbox5.addWidget(stat_op16, 3,2)
        vbox5.addWidget(stat_op17, 4,2)
        groupBox5.setLayout(vbox5)
        
        btn1 = QtGui.QPushButton("OK", self.option_win)
        btn1.clicked.connect(lambda: self.input_check(unicode(dir), unicode(job)))
        btn1.resize(btn1.minimumSizeHint())
        btn1.move(600,830)
        btn1.setToolTip("Finalize And Generate")
        
        btn2 = QtGui.QPushButton("Close", self.option_win)
        btn2.clicked.connect(self.clse2)
        btn2.resize(btn2.minimumSizeHint())
        btn2.move(700,830)
        btn2.setToolTip("Close Current Window")
        
        groupBox6 = QGroupBox("Ending Remarks", self.option_win)
        groupBox6.setGeometry(QRect(10, 610, 780, 180))
        self.groupBox6 = groupBox6
        vbox6 = QGridLayout()
        
        remark_label = QLabel()
        remark_label.setText("Remarks")
        remark_statement = QtGui.QLineEdit(self.option_win)
        
        remark_label2 = QLabel()
        remark_label2.setText("Project Manager")
        proj_man = QtGui.QLineEdit(self.option_win)
        proj_man.setText(def_proj_man)
        self.proj_man = proj_man
        
        vbox6.addWidget(remark_label, 0,0)
        vbox6.addWidget(remark_statement, 1,0)
        vbox6.addWidget(remark_label2, 2,0)
        vbox6.addWidget(proj_man, 3,0)
        groupBox6.setLayout(vbox6)
        
        self.option_win.installEventFilter(self)
        
        self.remark_statement = remark_statement
        self.proj_man = proj_man
        self.remark_label = remark_label
        self.remark_label2 = remark_label2
        
        self.option_win.show()
        
    def input_check(self, dir, job): #validation of user inputs
        #groupbox1
        
        proj = unicode(self.proj.text())
        loc = unicode(self.loc.text())
        to_whom = unicode(self.to_whom.toPlainText())
        err_check = 0
        err_string = ""
        
        if self.trans.text() and any(char.isalpha() or char.isdigit() for char in unicode(self.trans.text())):
            if not any(char.isalpha() for char in unicode(self.trans.text())):
                self.label1.setText("Transmittal #")
                self.label1.setStyleSheet('QLabel {color: black}')
            else:
                if err_string:
                    err_string += ", Numbers only in Transmittal #"
                else:
                    err_string += "Numbers only in Transmittal #"
                self.label1.setText("*Transmittal #")
                self.label1.setStyleSheet('QLabel {color: red}')
        else:
            if err_string:
                err_string += ", Transmittal # field is required"
            else:
                err_string += "Transmittal # field is required"
                
            self.label1.setText("*Transmittal #")
            self.label1.setStyleSheet('QLabel {color: red}')
            
        if self.dte.text() and any(char.isalpha() or char.isdigit() for char in unicode(self.dte.text())):
            self.label2.setText("Date")
            self.label2.setStyleSheet('QLabel {color: black}')
        else:
            if err_string:
                err_string += ", Date field is required"
            else:
                err_string += "Date field is required"
                
            self.label2.setText("*Date")
            self.label2.setStyleSheet('QLabel {color: red}')
            
        if self.proj_num.text() and any(char.isalpha() or char.isdigit() for char in unicode(self.proj_num.text())):
            if not any(char.isalpha() for char in unicode(self.proj_num.text())):
                self.label3.setText("Project #")
                self.label3.setStyleSheet('QLabel {color: black}')
            else:
                if err_string:
                    err_string += ", Numbers only in Project # field"
                else:
                    err_string += "Numbers only in Project # field"
                    
                self.label3.setText("*Project #")
                self.label3.setStyleSheet('QLabel {color: red}')
        else:
            if err_string:
                err_string += ", Project # field is required"
            else:
                err_string += "Project # field is required"
                
            self.label3.setText("*Project #")
            self.label3.setStyleSheet('QLabel {color: red}')
            
        if proj and any(char.isalpha() or char.isdigit() for char in proj):
            self.label4.setText("Project")
            self.label4.setStyleSheet('QLabel {color: black}')
        else:
            if err_string:
                err_string += ", Project field is required"
            else:
                err_string += "Project field is required"
                
            self.label4.setText("*Project")
            self.label4.setStyleSheet('QLabel {color: red}')
            
        if loc and any(char.isalpha() or char.isdigit() for char in loc):
            self.label5.setText("Location")
            self.label5.setStyleSheet('QLabel {color: black}')
        else:
            if err_string:
                err_string += ", Location field is required"
            else:
                err_string += "Location field is required"
            self.label5.setText("*Location")
            self.label5.setStyleSheet('QLabel {color: red}')
            
        if to_whom and any(char.isalpha() or char.isdigit() for char in to_whom):
            self.label6.setText("TO")
            self.label6.setStyleSheet('QLabel {color: black}')
        else:
            if err_string:
                err_string += ", TO field is required"
            else:
                err_string += "TO field is required"
            self.label6.setText("*TO")
            self.label6.setStyleSheet('QLabel {color: red}')
            
            
        #groupbox2
        op1 = self.op1
        op2 = self.op2
        desc = unicode(self.desc.text())
        if op1.isChecked() or op2.isChecked():
            self.op1.setText("Herewith")
            self.op2.setText("Under separate cover via: ")
            self.op1.setStyleSheet('QCheckBox {color: Black}')
            self.op2.setStyleSheet('QCheckBox {color: Black}')
            if op2.isChecked():
                if desc:
                    self.op2.setText("Under separate cover via: ")
                    self.op2.setStyleSheet('QCheckBox {color: Black}')
                else:
                    self.op2.setText("*Under separate cover via: ")
                    self.op2.setStyleSheet('QCheckBox {color: red}')
                    
                    if err_string:
                        err_string += '\n \n'
                        err_string += "With/Separate: Please also fill the text box beside 'Under separate cover via'"
                    else:
                        err_string += "With/Separate: Please also fill the text box beside 'Under separate cover via'"
            else:
                if desc:
                    self.op2.setText("*Under separate cover via: ")
                    self.op2.setStyleSheet('QCheckBox {color: red}')
                    
                    if err_string:
                        err_string += '\n \n'
                        err_string += "With/Separate: Please Check 'Under separate cover via'"
                    else:
                        err_string += "With/Separate: Please Check 'Under separate cover via'"
        else:
            self.op1.setText("*Herewith")
            self.op2.setText("*Under separate cover via: ")
            self.op1.setStyleSheet('QCheckBox {color: red}')
            self.op2.setStyleSheet('QCheckBox {color: red}')
            if err_string:
                err_string += '\n \n'
                err_string += "With/Separate: Please select a checkbox"
                
            else:
                err_string += "With/Separate: Please select a checkbox"
                
                
        #groupbox3
        stat_check = 0
        ckc_test = 0
        checkBoxList  = self.groupBox3.findChildren(QtGui.QCheckBox)
        for i in range(len(checkBoxList)):
            if checkBoxList[i].isChecked():
                stat_check = 1
                break
                
        if stat_check <= 0:
            for i in range(len(checkBoxList)):
                new_chktxt = checkBoxList[i].text()
                if '*' not in checkBoxList[i].text():
                    new_chktxt = '*' + checkBoxList[i].text()
                    
                checkBoxList[i].setText(new_chktxt)
                checkBoxList[i].setStyleSheet('QCheckBox {color: red}')
                if err_string: 
                    if ckc_test == 0:
                        err_string += "\n \n Files To Be Sent: Please select a checkbox"
                        ckc_test += 1
                else:
                    if ckc_test == 0:
                        err_string += "Files To Be Sent: Please select a checkbox"
                        ckc_test += 1
        else:
            for i in range(len(checkBoxList)):
                checkBoxList[i].setText(unicode(checkBoxList[i].text()).replace('*', ''))
                checkBoxList[i].setStyleSheet('QCheckBox {color: black}')
                
        #groupbox5
        stat_check = 0
        ckc_test = 0
        checkBoxList  = self.groupBox5.findChildren(QtGui.QCheckBox)
        for i in range(len(checkBoxList)):
            if checkBoxList[i].isChecked():
                stat_check = 1
                break
                
        if stat_check <= 0:
            for i in range(len(checkBoxList)):
                new_chktxt = checkBoxList[i].text()
                if '*' not in checkBoxList[i].text():
                    new_chktxt = '*' + checkBoxList[i].text()
                    
                checkBoxList[i].setText(new_chktxt)
                checkBoxList[i].setStyleSheet('QCheckBox {color: red}')
                if err_string: 
                    if ckc_test == 0:
                        err_string += "\n \n Status: Please select a checkbox"
                        ckc_test += 1
                else:
                    if ckc_test == 0:
                        err_string += "Status: Please select a checkbox"
                        ckc_test += 1
        else:
            for i in range(len(checkBoxList)):
                checkBoxList[i].setText(unicode(checkBoxList[i].text()).replace('*', ''))
                checkBoxList[i].setStyleSheet('QCheckBox {color: black}')
                
        #groupbox6
        proj_man = unicode(self.proj_man.text())
        remark_label2 = self.remark_label2
        
        if proj_man and any(char.isalpha() or char.isdigit() for char in proj_man):
            self.remark_label2.setText("Project Manager")
            self.remark_label2.setStyleSheet('QLabel {color: black}')
        else:
            if err_string:
                err_string += "\n \n Ending Remarks: Project Manager field is required"
            else:
                err_string += "Ending Remarks: Project Manager field is required"
                
            self.remark_label2.setText("*Project Manager")
            self.remark_label2.setStyleSheet('QLabel {color: red}')
            
        if err_string:
            self.warnmsg(err_string)
            
        else:
            self.trans_fetch(dir, job)
            
    def trans_fetch(self, trasmittal_dir, job):#fetches all the files in the FOR SENDING of the choosen transmittal
        file_det = {}
        for fol in os.listdir(trasmittal_dir):
            for dir, folders, files in os.walk(os.path.join(trasmittal_dir, fol)):
                new_dir = dir.split('\\')
                if new_dir[len(new_dir)-1] and files:
                    file_det[new_dir[len(new_dir)-1]] = files
                    
        
        output_settings = {}
        output_settings2 = {}
        output_settings3 = {}
        output_settings4 = {}
        output_settings5 = {}
        
        output_settings['Transmittal #'] = self.trans.text()
        output_settings['Date'] = self.dte.text()
        output_settings['Project #'] = self.proj_num.text()
        output_settings['Project'] = self.proj.text()
        output_settings['Location'] = self.loc.text()
        self.con_attn = ''
        
        if self.attn.text() and any(char.isalpha() or char.isdigit() for char in unicode(self.attn.text())):
            self.con_attn = unicode(self.attn.text()).replace('&','&amp;')
        
        
        if self.desc.text() and any(char.isalpha() or char.isdigit() for char in unicode(self.desc.text())):
            output_settings2['desc'] = unicode(self.desc.text()).replace('&','&amp;')
            
        if self.op2.isChecked():
            output_settings2['Under Separate Cover Via:'] = 'X'
        else:
            output_settings2['Under Separate Cover Via:'] = ' '
            
        if self.op1.isChecked():
            output_settings2['HereWith'] = 'X'
        else:
            output_settings2['HereWith'] = ' '
            
            
        checkBoxList  = self.groupBox3.findChildren(QtGui.QCheckBox)
        for i in range(len(checkBoxList)):
            if checkBoxList[i].isChecked():
                output_settings3[unicode(checkBoxList[i].text())] = 'X'
            else:
                output_settings3[unicode(checkBoxList[i].text())] = ' '
                
        checkBoxList2  = self.groupBox5.findChildren(QtGui.QCheckBox)
        for i in range(len(checkBoxList2)):
            if checkBoxList2[i].isChecked():
                output_settings4[unicode(checkBoxList2[i].text())] = 'X'
            else:
                output_settings4[unicode(checkBoxList2[i].text())] = ' '
                
        output_settings5['Project Manager'] = unicode(self.proj_man.text()).replace('&','&amp;')
        
        if unicode(self.remark_statement.text()) and any(char.isalpha() or char.isdigit() for char in unicode(self.remark_statement.text())):
            output_settings5['Remarks'] = unicode(self.remark_statement.text()).replace('&','&amp;')
        
        char_det2 = 0
        
        file_name = file_det
        self.output_settings = output_settings
        self.output_settings2 = output_settings2
        self.output_settings3 = output_settings3
        self.output_settings4 = output_settings4
        self.output_settings5 = output_settings5
        self.confirmation(file_name, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job)
        
    def confirmation(self, file_name, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job): #filters and displays a table to edit what files would the user like to include based on the File Type Settings
        
        option_tbl = option_win(self)
        option_tbl.resize(800, 865)
        option_tbl.setFixedSize(option_tbl.size())
        option_tbl.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        option_tbl.setWindowTitle("Files Detected")
        option_tbl.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        
        table = QTableWidget()
        table.resize(400, 300)
        
        header_text = []
        row_len = 0
        for desc, file in file_name.iteritems():
            header_text += [desc]
            if len(file) > row_len:
                row_len = len(file)
                
        filter_sheets = QLineEdit(option_tbl)
        filter_sheets.setPlaceholderText("Strict Filter")
        filter_sheets.setToolTip("Will Search Strictly")
        filter_sheets.move(330,35)
        filter_sheets.textChanged.connect(lambda: self.filter_match(table, filter_sheets, row_len, file_name))
        
        table.setRowCount(row_len)    
        table.setColumnCount(len(header_text))
        table.setHorizontalHeaderLabels((header_text))
        
        btn1 = QtGui.QPushButton("OK", option_tbl)
        btn1.clicked.connect(lambda: self.fetch_input(table, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl))
        btn1.resize(btn1.minimumSizeHint())
        btn1.move(600,830)
        
        btn2 = QtGui.QPushButton("Cancel", option_tbl)
        btn2.clicked.connect(lambda: self.table_close(option_tbl))
        btn2.resize(btn2.minimumSizeHint())
        btn2.move(700,830)
        
        groupBox1 = QGroupBox("", option_tbl)
        groupBox1.setGeometry(QRect(10, 10, 780, 155))
        vbox = QGridLayout()
        
        vbox.addWidget(filter_sheets, 0,0)
        groupBox1.setLayout(vbox)
        
        groupBox2 = QGroupBox("", option_tbl)
        groupBox2.setGeometry(QRect(10, 115, 780, 710))
        vbox2 = QGridLayout()
        
        vbox2.addWidget(table, 0,0)
        groupBox2.setLayout(vbox2)
        
        option_tbl.setWindowModality(QtCore.Qt.ApplicationModal)
        option_tbl.show()
        self.table_generation(file_name, table)
        
    def table_generation(self, file_name, table):
        asd = table.horizontalHeader().model()
        
        for index in range(asd.columnCount()):
            counter = 0
            for desc, file in file_name.iteritems():
                if desc == unicode(table.horizontalHeaderItem(index).text()):
                    file.sort(key = alphanum_key)
                    for items in file:
                        if '.db' not in os.path.splitext(items)[1] and '.png' not in os.path.splitext(items)[1] and '.bmp' not in os.path.splitext(items)[1]:
                            chkBoxItem = QtGui.QTableWidgetItem(items)
                            chkBoxItem.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                            chkBoxItem.setCheckState(QtCore.Qt.Checked)
                            table.setItem(counter, index, chkBoxItem)
                            counter += 1
        table.resizeColumnsToContents()
        
    def fetch_input(self, table, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl): #stores and readies the files for segregation and concatination
        sheet_matcher = {}
        
        for i in range (table.columnCount()):
            for j in range(table.rowCount()):
                if table.item(j,i) is not None and table.item(j,i).checkState():
                    if unicode(table.horizontalHeaderItem(i).text()) in sheet_matcher:
                        sheet_matcher[unicode(table.horizontalHeaderItem(i).text())].append(unicode(table.item(j,i).text()))
                    else:
                        sheet_matcher[unicode(table.horizontalHeaderItem(i).text())] = [unicode(table.item(j,i).text())]
        self.generation_check(sheet_matcher, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl)
        
    def generation_check(self, sheet_matcher, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl):
        if sheet_matcher:
            precon_sheets = {}
            miscellaneous_sheets = {}
            rev_match = {}
            for desc, fold in sheet_matcher.iteritems():
                added_sheet = []
                for file in fold:
                    full_sheet = ''
                    sheet_key = file
                    sheet_rev = ''
                    if '.7z' == os.path.splitext(file)[1] or'.xsr' == os.path.splitext(file)[1] or '.kss' == os.path.splitext(file)[1] or '.ifc' == os.path.splitext(file)[1] or '.xls' == os.path.splitext(file)[1] or '.xlsx' == os.path.splitext(file)[1]:
                        if file not in added_sheet:
                            if desc in miscellaneous_sheets:
                                miscellaneous_sheets[desc].append(file)
                                added_sheet += [file]
                            else:
                                miscellaneous_sheets[desc] = [file]
                                added_sheet += [file]
                    else:
                        sheet_key = os.path.splitext(file)[0]
                        special_char = '!@#$%^&*()_+=-{}[]|"\'?/><, '
                        rev = 'rev'
                        crev = 'Rev'
                        acrev = 'REV'
                        revision = 'revision'
                        crevision = 'Revision'
                        acrevision = 'REVISION'
                        sheet_extension = os.path.splitext(file)[1]
                        
                        for char in sheet_key:
                            if char in special_char:
                                sheet_key = sheet_key.replace(char, '_')
                        sheet_key = sheet_key.replace(revision, '_')
                        sheet_key = sheet_key.replace(crevision, '_')
                        sheet_key = sheet_key.replace(acrevision, '_')
                        
                        sheet_key = sheet_key.replace(rev, '_')
                        sheet_key = sheet_key.replace(crev, '_')
                        sheet_key = sheet_key.replace(acrev, '_')
                        temp_string = ''
                        
                        for i in sheet_key.split('_'):
                            if not temp_string and i:
                                temp_string += i
                            else:
                                if i:
                                    temp_string += '_' + i
                                    
                        sheet_key = temp_string
                        
                        if sheet_key.count('_') == 2:
                            sheet_part = sheet_key.split('_')
                            sheet_name = ''
                            sheet_rev = sheet_key.split('_')[2]
                            for item in range(len(sheet_part)):
                                has_dig = 0
                                has_str = 0
                                for char in sheet_part[item]:
                                    if char.isdigit():
                                        has_dig = 1
                                    if char.isalpha():
                                        has_str = 1
                                        
                                if sheet_name:
                                    if len(sheet_name) < sheet_part[item]:
                                        if has_dig and has_str:
                                            sheet_name = sheet_part[item]
                                else:
                                    if has_dig and has_str:
                                        sheet_name = sheet_part[item]
                                        
                        if sheet_key.count('_') == 1:
                            sheet_name = sheet_key.split('_')[0]
                            sheet_rev = sheet_key.split('_')[1]
                            
                        if sheet_key.count('_') == 0:
                            sheet_name = os.path.splitext(file)[0]
                        #Check if necessarry    
                        if 'nc1' in sheet_extension:
                            checker_list_items = []
                            list_item = ''
                            string_switch = False
                            string_to_check = sheet_name + '_' + sheet_rev
                            for char in string_to_check:
                                if string_switch:
                                    checker_list_items += [list_item]
                                    string_switch = False
                                    list_item = ''
                                if char.isalpha():
                                    list_item += char
                                elif len(list_item) > 0:
                                    string_switch = True
                            if len(checker_list_items) > 1:
                                patt_check1 = '@#@#'
                                patt_check2 = '@#_@#'
                                sheet_pattern = ''
                                for char in string_to_check:
                                    if char.isdigit():
                                        sheet_pattern += '#'
                                    elif char.isalpha():
                                        sheet_pattern += '@'
                                    else:
                                        sheet_pattern += char
                                #include no key
                                if patt_check1 in sheet_pattern:
                                    pat_act = 0
                                    string_pat = ''
                                    new_sheet_string = ''
                                    for char in range(len(sheet_name)):
                                        if sheet_name[char].isalpha() and char > 0:
                                            pat_act += 1
                                        if pat_act > 0:
                                            string_pat += sheet_name[char]
                                    if len(sheet_rev) > 0:
                                        new_sheet_string = sheet_name.replace(string_pat, '') + '_' + string_pat + '^' + sheet_rev
                                    else:
                                        new_sheet_string = sheet_name.replace(string_pat, '') + '_' + string_pat + '^'
                                    sheet_key =  new_sheet_string
                                if patt_check2 in sheet_pattern:
                                    pat_act = 0
                                    string_pat = ''
                                    new_sheet_string = ''
                                    for char in range(len(sheet_name)):
                                        if sheet_name[char].isalpha() and char > 0:
                                            pat_act += 1
                                        if pat_act > 0:
                                            string_pat += sheet_name[char]
                                    if len(sheet_rev) > 0:
                                        new_sheet_string = sheet_name.replace(string_pat, '') + '_' + string_pat + '*' + sheet_rev
                                    else:
                                        new_sheet_string = sheet_name.replace(string_pat, '') + '_' + string_pat + '*'
                                    sheet_key =  new_sheet_string
                                    
                        for nme in sheet_key.split('_'):
                            if not full_sheet and nme:
                                full_sheet += nme
                            else:
                                if nme:
                                    full_sheet += '_' + nme
                                    
                        full_sheet += sheet_extension
                        if sheet_key not in added_sheet:
                            if desc in precon_sheets:
                                precon_sheets[desc].append(full_sheet)
                                added_sheet += [sheet_key]
                            else:
                                precon_sheets[desc] = [full_sheet]
                                added_sheet += [sheet_key]
                                
            for rev_desc, rev_sheets in precon_sheets.iteritems():
                sheet_name = ''
                sheet_revision = ''
                for rev_items in rev_sheets:
                    sheet_name = os.path.splitext(rev_items)[0]
                    sheet_ext = os.path.splitext(rev_items)[1]
                    
                    if sheet_name.count('_') == 2:
                        sheet_revision = sheet_name.split('_')[2]
                        
                        if len(sheet_revision) > 1 and '*' not in sheet_revision and '^' not in sheet_revision:
                            new_rev = ''
                            for char in sheet_revision:
                                if char.isdigit():
                                    new_rev += char
                            sheet_revision = new_rev
                            
                    if sheet_name.count('_') == 1:
                        sheet_revision = sheet_name.split('_')[1]
                        
                        if len(sheet_revision) > 1  and '*' not in sheet_revision and '^' not in sheet_revision:
                            new_rev = ''
                            for char in sheet_revision:
                                if char.isdigit():
                                    new_rev += char
                            sheet_revision = new_rev
                            
                    if sheet_name.count('_') == 0:
                        sheet_revision = 'A'
                        
                    if sheet_revision not in rev_match:
                        rev_match[sheet_revision] = [rev_items]
                    else:
                        rev_match[sheet_revision].append(rev_items) 
                        
        self.file_comp(precon_sheets, miscellaneous_sheets, rev_match, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl)
        
    def file_comp(self, precon_sheets, miscellaneous_sheets, rev_match, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl):
        context = {}
        matches = {}
        rev_len = 0
        for rev, sheet_in_rev in rev_match.iteritems():
            sheet_in_rev.sort(key = alphanum_key)
            rev_len = len(sheet_in_rev)-1
            for desc, pre_processed_sheets in precon_sheets.iteritems():
                pre_processed_sheets.sort(key = alphanum_key)
                sorted_sheets = pre_processed_sheets
                upper_list = []
                list_map = []
                sheet_check_1 = ''
                sheet_check_2 = ''
                counter = 0
                for sheet_index in range(len(sorted_sheets)):
                    sheet = sorted_sheets[sheet_index]
                    if sheet in sheet_in_rev:
                        if len(sheet_in_rev) > 2:
                            main_sheet = ''
                            sheet_key = ''
                            sheet_number = ''
                            new_key = ''
                            
                            if sheet.count('_') == 2:
                                split_sheet = os.path.splitext(sheet)[0]
                                sheet_part = split_sheet.split('_')
                                for item in range(len(sheet_part)):
                                    has_dig = 0
                                    has_str = 0
                                    for char in sheet_part[item]:
                                        if char.isdigit():
                                            has_dig = 1
                                        if char.isalpha():
                                            has_str = 1
                                            
                                    if main_sheet:
                                        if len(main_sheet) < sheet_part[item]:
                                            if has_dig and has_str:
                                                main_sheet = sheet_part[item]
                                    else:
                                        if has_dig and has_str:
                                            main_sheet = sheet_part[item]
                                            
                                for char in main_sheet:
                                    if char.isdigit():
                                        sheet_number += char
                                    elif char.isalpha():
                                        sheet_key += char
                                        
                            elif sheet.count('_') == 1:
                                split_sheet = os.path.splitext(sheet)[0]
                                main_sheet = split_sheet.split('_')[0]
                                for char in main_sheet:
                                    if char.isdigit():
                                        sheet_number += char
                                    elif char.isalpha():
                                        sheet_key += char
                                        
                            elif sheet.count('_') == 0:
                                split_sheet = os.path.splitext(sheet)[0]
                                main_sheet = split_sheet
                                for char in main_sheet:
                                    if char.isdigit():
                                        sheet_number += char
                                    elif char.isalpha():
                                        sheet_key += char
                                        
                            if sheet_check_1:
                                sheet_check_2 = sheet_number
                                if prev_key == sheet_key:
                                    if (int(sheet_check_2) - int(sheet_check_1)) != counter+1:
                                        if len(pre_processed_sheets)-1 == sheet_index:
                                            if counter > 2:
                                                index_1 = ''
                                                index_2 = ''
                                                if '^' in rev:
                                                    if rev[len(rev)-1] == '^':
                                                        index_1 = list_map[counter - counter] + rev.replace('^','')
                                                        index_2 = list_map[counter] + rev.replace('^','')
                                                    else:
                                                        index_1 = list_map[counter - counter] + rev.replace('_','').replace('^','_')
                                                        index_2 = list_map[counter] + rev.replace('_','').replace('^','_')
                                                elif '*' in rev:
                                                    if rev[len(rev)-1] == '*':
                                                        index_1 = list_map[counter - counter] + rev.replace('*','')
                                                        index_2 = list_map[counter] + rev.replace('*','')
                                                    else:
                                                        index_1 = list_map[counter - counter] + rev.replace('_','').replace('*','_')
                                                        index_2 = list_map[counter] + rev.replace('_','').replace('*','_')
                                                else:
                                                    
                                                    index_1 = list_map[counter - counter]
                                                    index_2 = list_map[counter]
                                                    
                                                string_concat = index_1 + ' thru ' + index_2
                                                upper_list += [string_concat]
                                            else:
                                                
                                                list_map += [main_sheet]
                                                upper_list += list_map
                                        else:
                                            if counter > 2:
                                                index_1 = ''
                                                index_2 = ''
                                                if '^' in rev:
                                                    if rev[len(rev)-1] == '^':
                                                        index_1 = list_map[counter - counter] + rev.replace('^','')
                                                        index_2 = list_map[counter] + rev.replace('^','')
                                                    else:
                                                        index_1 = list_map[counter - counter] + rev.replace('_','').replace('^','_')
                                                        index_2 = list_map[counter] + rev.replace('_','').replace('^','_')
                                                        
                                                elif '*' in rev:
                                                    if rev[len(rev)-1] == '*':
                                                        index_1 = list_map[counter - counter] + rev.replace('*','')
                                                        index_2 = list_map[counter] + rev.replace('*','')
                                                    else:
                                                        index_1 = list_map[counter - counter] + rev.replace('_','').replace('*','_')
                                                        index_2 = list_map[counter] + rev.replace('_','').replace('*','_')
                                                else:
                                                    index_1 = list_map[counter - counter]
                                                    index_2 = list_map[counter]
                                                    
                                                string_concat = index_1 + ' thru ' + index_2
                                                upper_list += [string_concat]
                                                list_map = [main_sheet]
                                                counter = 0
                                                sheet_check_1 = sheet_check_2
                                                sheet_check_2 = ''
                                                prev_key = sheet_key
                                            else:
                                                upper_list += list_map
                                                list_map = [main_sheet]
                                                counter = 0
                                                sheet_check_1 = sheet_check_2
                                                sheet_check_2 = ''
                                                prev_key = sheet_key
                                    else:
                                        if len(pre_processed_sheets)-1 == sheet_index:
                                            list_map += [main_sheet]
                                            
                                            counter += 1 
                                            if counter > 2:
                                                index_1 = ''
                                                index_2 = ''
                                                if '^' in rev:
                                                    if rev[len(rev)-1] == '^':
                                                        index_1 = list_map[counter - counter] + rev.replace('^','')
                                                        index_2 = list_map[counter] + rev.replace('^','')
                                                    else:
                                                        index_1 = list_map[counter - counter] + rev.replace('_','').replace('^','_')
                                                        index_2 = list_map[counter] + rev.replace('_','').replace('^','_')
                                                        
                                                elif '*' in rev:
                                                    if rev[len(rev)-1] == '*':
                                                        index_1 = list_map[counter - counter] + rev.replace('*','')
                                                        index_2 = list_map[counter] + rev.replace('*','')
                                                    else:
                                                        index_1 = list_map[counter - counter] + rev.replace('_','').replace('*','_')
                                                        index_2 = list_map[counter] + rev.replace('_','').replace('*','_')
                                                        
                                                else:
                                                    index_1 = list_map[counter - counter]
                                                    index_2 = list_map[counter]
                                                    
                                                string_concat = index_1 + ' thru ' + index_2
                                                upper_list += [string_concat]
                                            else:
                                                upper_list += list_map
                                        else:
                                            list_map += [main_sheet]
                                            counter += 1 
                                            if rev_len == sheet_in_rev.index(sheet):
                                                if counter > 2:
                                                    index_1 = ''
                                                    index_2 = ''
                                                    if '^' in rev:
                                                        if rev[len(rev)-1] == '^':
                                                            index_1 = list_map[counter - counter] + rev.replace('^','')
                                                            index_2 = list_map[counter] + rev.replace('^','')
                                                        else:
                                                            index_1 = list_map[counter - counter] + rev.replace('_','').replace('^','_')
                                                            index_2 = list_map[counter] + rev.replace('_','').replace('^','_')
                                                    
                                                    elif '*' in rev == '*':
                                                        if rev[len(rev)-1] == '*':
                                                            index_1 = list_map[counter - counter] + rev.replace('*','')
                                                            index_2 = list_map[counter] + rev.replace('*','')
                                                        else:
                                                            index_1 = list_map[counter - counter] + rev.replace('_','').replace('*','_')
                                                            index_2 = list_map[counter] + rev.replace('_','').replace('*','_')
                                                        
                                                    else:
                                                        index_1 = list_map[counter - counter]
                                                        index_2 = list_map[counter]
                                                        
                                                    string_concat = index_1 + ' thru ' + index_2
                                                    upper_list += [string_concat]
                                                else:
                                                    upper_list += list_map
                                                    
                                else:
                                    if len(pre_processed_sheets)-1 == sheet_index:
                                        list_map += [main_sheet]
                                        counter += 1 
                                        if counter > 2:
                                            index_1 = ''
                                            index_2 = ''
                                            if '^' in rev:
                                                if rev[len(rev)-1] == '^':
                                                    index_1 = list_map[counter - counter] + rev.replace('^','')
                                                    index_2 = list_map[counter] + rev.replace('^','')
                                                else:
                                                    index_1 = list_map[counter - counter] + rev.replace('_','').replace('^','_')
                                                    index_2 = list_map[counter] + rev.replace('_','').replace('^','_')
                                                    
                                            elif '*' in rev:
                                                if rev[len(rev)-1] == '*':
                                                    index_1 = list_map[counter - counter] + rev.replace('*','')
                                                    index_2 = list_map[counter] + rev.replace('*','')
                                                else:
                                                    index_1 = list_map[counter - counter] + rev.replace('_','').replace('*','_')
                                                    index_2 = list_map[counter] + rev.replace('_','').replace('*','_')
                                            else:
                                                index_1 = list_map[counter - counter]
                                                index_2 = list_map[counter]
                                                
                                            string_concat = index_1 + ' thru ' + index_2
                                            upper_list += [string_concat]
                                        else:
                                            upper_list += list_map
                                    else:
                                        if counter > 2:
                                            index_1 = ''
                                            index_2 = ''
                                            if '^' in rev:
                                                if rev[len(rev)-1] == '^':
                                                    index_1 = list_map[counter - counter] + rev.replace('^','')
                                                    index_2 = list_map[counter] + rev.replace('^','')
                                                else:
                                                    index_1 = list_map[counter - counter] + rev.replace('_','').replace('^','_')
                                                    index_2 = list_map[counter] + rev.replace('_','').replace('^','_')
                                                    
                                            elif '*' in rev:
                                                if rev[len(rev)-1] == '*':
                                                    index_1 = list_map[counter - counter] + rev.replace('*','')
                                                    index_2 = list_map[counter] + rev.replace('*','')
                                                else:
                                                    index_1 = list_map[counter - counter] + rev.replace('_','').replace('*','_')
                                                    index_2 = list_map[counter] + rev.replace('_','').replace('*','_')
                                                    
                                            else:
                                                index_1 = list_map[counter - counter]
                                                index_2 = list_map[counter]
                                                
                                            string_concat = index_1 + ' thru ' + index_2
                                            upper_list += [string_concat]
                                            list_map = [main_sheet]
                                            counter = 0
                                            sheet_check_1 = sheet_check_2
                                            sheet_check_2 = ''
                                            prev_key = sheet_key
                                        else:
                                            upper_list += list_map
                                            list_map = [main_sheet]
                                            counter = 0
                                            sheet_check_1 = sheet_check_2
                                            sheet_check_2 = ''
                                            prev_key = sheet_key
                            else:
                                sheet_check_1 = sheet_number
                                list_map += [main_sheet]
                                prev_key = sheet_key
                                if len(pre_processed_sheets)-1 == sheet_index:
                                    upper_list += list_map
                        else:
                            main_sheet = ''
                            sheet_key = ''
                            sheet_number = ''
                            new_key = ''
                            if sheet.count('_') == 2:
                                split_sheet = os.path.splitext(sheet)[0]
                                sheet_part = split_sheet.split('_')
                                for item in range(len(sheet_part)):
                                    has_dig = 0
                                    has_str = 0
                                    for char in sheet_part[item]:
                                        if char.isdigit():
                                            has_dig = 1
                                        if char.isalpha():
                                            has_str = 1
                                            
                                    if main_sheet:
                                        if len(main_sheet) < sheet_part[item]:
                                            if has_dig and has_str:
                                                main_sheet = sheet_part[item]
                                    else:
                                        if has_dig and has_str:
                                            main_sheet = sheet_part[item]
                                            
                                for char in main_sheet:
                                    if char.isdigit():
                                        sheet_number += char
                                    elif char.isalpha():
                                        sheet_key += char
                                        
                            elif sheet.count('_') == 1:
                                split_sheet = os.path.splitext(sheet)[0]
                                main_sheet = split_sheet.split('_')[0]
                                for char in main_sheet:
                                    if char.isdigit():
                                        sheet_number += char
                                    elif char.isalpha():
                                        sheet_key += char
                                        
                            elif sheet.count('_') == 0:
                                split_sheet = os.path.splitext(sheet)[0]
                                main_sheet = split_sheet
                                for char in main_sheet:
                                    if char.isdigit():
                                        sheet_number += char
                                    elif char.isalpha():
                                        sheet_key += char
                            upper_list += [main_sheet]
                            
                post_concat_filter = []
                for items in upper_list:
                    if items:
                        if 'thru' not in items:
                            if '^' in rev:
                                if rev[len(rev)-1] == '^':
                                    post_concat_filter += [items + rev.replace('^','')]
                                else:
                                    get_rev = rev.split('^')[0]
                                    post_concat_filter += [items + get_rev]
                            elif '*' in rev:
                                post_concat_filter += [items + rev.replace('*','_')]
                            else:
                                post_concat_filter += [items]
                                
                        else:
                            if '*' in rev:
                                temp_sheet = items.split(' ')
                                temp_sheet[0] = temp_sheet[0] + rev.replace('*','_')
                                temp_sheet[2] = temp_sheet[2] + rev.replace('*','_')
                                temp_concat = ''
                                for temp_item in temp_sheet:
                                    if not temp_concat:
                                        temp_concat += temp_item
                                    else:
                                        temp_concat += ' ' + temp_item
                                post_concat_filter += [temp_concat]
                            else:
                                post_concat_filter += [items]
                if post_concat_filter:
                    post_rev = ''
                    if '^' in rev:
                        if rev.split('^')[1]:
                            post_rev = rev.split('^')[1]
                        else:
                            post_rev = rev.replace('^', '')
                    elif '*' in rev:
                        post_rev = ''
                    else:
                        post_rev = rev
                        
                    if post_rev not in matches:
                        matches[post_rev] = {desc:post_concat_filter}
                    else:
                        if desc in matches[post_rev]:
                            matches[post_rev][desc].extend(post_concat_filter)
                        else:
                            matches[post_rev][desc] = post_concat_filter
                            
        for misc_desc, misc_sheet in miscellaneous_sheets.iteritems():
            if 'norev' not in context:
                context['norev'] = {misc_desc:misc_sheet}
            else:
                if misc_desc not in context['norev']:
                    context['norev'][misc_desc] = misc_sheet
                else:
                    context['norev'][misc_desc].extend(misc_sheet)   
                    
        main_list = []
        for items in matches:
            if 'norev' not in items:
                for desc,sheets in matches[items].iteritems():
                    main_list = []
                    string_list = ''
                    string_counter = 0
                    for file in range(len(sheets)):
                        if sheets[file]:
                            if 'thru' in sheets[file]:
                                string_counter += 5
                                if len(sheets) > 1 and file < len(sheets)-1:
                                    string_list += sheets[file] + ', '
                                if file == len(sheets)-1:
                                    main_list += [sheets[file]]
                            else:
                                if len(sheets) > 1 and file < len(sheets)-1:
                                    string_list += sheets[file] + ', '
                                if file == len(sheets)-1:
                                    string_list += sheets[file]
                                string_counter += 1
                                
                            if string_counter >= 6:
                                if ', ' == string_list[len(string_list)-2:]:
                                    string_list = string_list[:len(string_list)-2]
                                main_list += [string_list]
                                string_list = ''
                                string_counter = 0
                                string_list = ''
                            else:
                                if file == len(sheets)-1:
                                    main_list += [string_list]            
                    if main_list:
                        if items not in context:
                            context[items] = {string.capwords(desc):main_list}
                        else:
                            if desc in context[items]:
                                context[items][string.capwords(desc)].extend(main_list)
                            else:
                                context[items][string.capwords(desc)] = main_list
                                
            else:
                for desc,sheets in matches[items].iteritems():
                    if items not in context:
                        context[items] = {string.capwords(desc):sheets}
                    else:
                        if desc in context[items]:
                            context[items][string.capwords(desc)].update(sheets)
                        else:
                            context[items][string.capwords(desc)] = sheets
                            
        if context:
            self.act_gen_word(context, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl)
            
    def act_gen_word(self, context, output_settings, output_settings2, output_settings3, output_settings4, output_settings5, job, option_tbl):
        for out_item in output_settings:
            if '&' in output_settings[out_item]:
                output_settings[out_item] = unicode(output_settings[out_item]).replace('&','&amp;')
            else:
                output_settings[out_item] = unicode(output_settings[out_item])
        self.option_tbl = option_tbl
        self.gen_done = 0
        string_counter = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        new_context = {}
        data_count = 0
        page_count = 0
        string_index = 0
        file_check = 0
        total_files = 0
        doc_name = unicode(output_settings['Project #']) + '_TRANS#' + unicode(output_settings['Transmittal #'])
        file_name = ''
        if os.path.exists(os.path.join('Transmittal Generated', job)):
            file_name = QtGui.QFileDialog.getSaveFileName(option_tbl, 'Save', os.path.join('Transmittal Generated', job, doc_name))
        else:
            os.makedirs(os.path.join('Transmittal Generated', job))
            file_name = QtGui.QFileDialog.getSaveFileName(option_tbl, 'Save', os.path.join('Transmittal Generated', job, doc_name))
            
        if file_name:
            for rev in context:
                for desc, sheets in context[rev].iteritems():
                    for file in sheets:
                        if len(file) > 0:
                            data_count += 1
                            total_files += 1
                            if data_count >= 13:
                                page_count += 1
                                data_count = 0
                                
            if page_count > 0:
                data_count = 0
                for rev in context:
                    for desc, sheets in context[rev].iteritems():
                        new_list = []
                        for file in sheets:
                            if len(file) > 0:
                                file_check += 1
                                data_count += 1
                                new_list += [file]
                                if data_count >= 13:
                                    if rev in new_context:
                                        new_context[rev.replace('\n', '')].update({ desc:new_list })
                                    else:
                                        new_context[rev.replace('\n', '')] = { desc:new_list }
                                    
                                    self.generate_word('Docs_generation',str(file_name + string_counter[string_index] + '.docx'), os.path.abspath('media'), 
                                        {
                                            'context':new_context, 'header_set': output_settings, 'pcount': string_counter[string_index],
                                            'send_type':output_settings2, 'sheet_type': output_settings3,
                                            'purpose':output_settings4, 'to':R(unicode(self.to_whom.toPlainText())),
                                            'attn':self.con_attn, 'end_remarks': output_settings5
                                        })
                                    new_context = {}
                                    new_list = []
                                    string_index += 1
                                    data_count = 0
                                else:
                                    if file_check == total_files:
                                        row_add = 13 - data_count
                                        dummy_list = []
                                        if row_add > 0:
                                            for i in range(row_add):
                                                dummy_list += ['none' + unicode(i)]
                                                
                                        if 'none' in new_context:
                                            new_context['none'].update({ 'none':dummy_list })
                                        else:
                                            new_context['none'] = { 'none':dummy_list }
                                            
                                        if rev in new_context:
                                            new_context[rev.replace('\n', '')].update({ desc:new_list })
                                        else:
                                            new_context[rev.replace('\n', '')] = { desc:new_list }
                                        self.gen_done = 1
                                        self.generate_word('Docs_generation',str(file_name + string_counter[string_index] + '.docx'), os.path.abspath('media'), 
                                            {
                                                'context':new_context, 'header_set': output_settings, 'pcount': string_counter[string_index],
                                                'send_type':output_settings2, 'sheet_type': output_settings3,
                                                'purpose':output_settings4, 'to':R(unicode(self.to_whom.toPlainText())),
                                                'attn':self.con_attn, 'end_remarks': output_settings5
                                            })
                                        
                        if rev in new_context:
                            new_context[rev.replace('\n', '')].update({ desc:new_list })
                        else:
                            new_context[rev.replace('\n', '')] = { desc:new_list }
                            
            else:
                sheet_count = 0
                total_row = 0
                dummy_list = []
                
                for rev in context:
                    for desc,sheets in context[rev].iteritems():
                        sheet_count += len(sheets)
                total_row = 13 - sheet_count
                if total_row > 0:
                    for extra_row in range(total_row):
                        dummy_list += ['none' + unicode(extra_row)]
                if dummy_list:
                    if 'none' in context:
                        context['none'].update({ 'none':dummy_list })
                    else:
                        context['none'] = { 'none':dummy_list }
                        output_settings['Project #']
                self.gen_done = 1
                self.generate_word('Docs_generation',str(file_name) + '.docx', os.path.abspath('media'), 
                    {
                        'context':context, 'header_set': output_settings, 'pcount': '',
                        'send_type':output_settings2, 'sheet_type': output_settings3,
                        'purpose':output_settings4, 'to':R(unicode(self.to_whom.toPlainText())),
                        'attn':self.con_attn, 'end_remarks': output_settings5
                    })
                    
            trans_num_def = output_settings['Transmittal #']
            
            trans_num_def = unicode(int(trans_num_def) + 1)
            
            if len(trans_num_def) < 4:
                string_len = 4 - len(trans_num_def)
                for i in range(string_len):
                    trans_num_def = '0' + trans_num_def
                    
            def_ini = ConfigParser.ConfigParser()
            def_ini.read(os.path.join('job_defaults', job, job + '.ini'))
            def_ini.set('defaults', 'Transmittal#', trans_num_def)
            with open(os.path.join('job_defaults', job, job + '.ini'), 'wb') as configfile:
                def_ini.write(configfile)
                
            
    def generate_word(self, template, filename, media_path=None, template_context=None):
        try:
            if template_context is None:
                template_context = {}
            jinja_env2  = Environment()
            doc = DocxTemplate(os.path.join("templates", "Word_Template", str(template) + '.docx'))
            doc.render(template_context,jinja_env2)
            doc.save(str(filename))
            if self.gen_done == 1:
                self.option_tbl.accept()
                self.option_win.accept()
                self.infmsg('Docx Generation Complete','File is Ready')
                self.gen_done = 0
        except Exception as e:
            self.warnmsg(str(e))
            
    #filter fields, event filters, and close command
    def filter_match(self, table, filter_sheets, row_len, file_name):
        if filter_sheets.text():
            table.setRowCount(0)
            table.setRowCount(row_len)
            asd = table.horizontalHeader().model()
            
            for index in range(asd.columnCount()):
                counter = 0
                for desc, file in file_name.iteritems():
                    if desc == str(table.horizontalHeaderItem(index).text()):
                        for items in file:
                            if '.db' not in os.path.splitext(items)[1] and '.png' not in os.path.splitext(items)[1] and '.bmp' not in os.path.splitext(items)[1] and str(filter_sheets.text()).lower() in items.lower():
                                chkBoxItem = QtGui.QTableWidgetItem(items)
                                chkBoxItem.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                                chkBoxItem.setCheckState(QtCore.Qt.Checked)
                                table.setItem(counter, index, chkBoxItem)
                                counter += 1
                                
            table.resizeColumnsToContents()
        else:
            table.setRowCount(row_len)
            self.table_generation(file_name, table)
            
    def action_change(self, job_name):
        job_match =  unicode(self.filter_field.text())
        if job_match:
            filter_job = []
            self.combo.clear()
            for items in job_name:
                if job_match in items.lower():
                    filter_job += [items]
                    
            self.combo.addItems(filter_job)
        else:
            job_name.sort(reverse = True)
            self.combo.addItems(job_name)
            
    def eventFilter(self, source, event):
        if event.type() == QtCore.QEvent.EnterWhatsThisMode:
            QtGui.QWhatsThis.leaveWhatsThisMode()
            os.startfile('Docs\how.html')
            return True
        else:
            return super(Window,self).eventFilter(source, event)
            
    def close_application(self):
        self.close()
        
    def open_about(self):
        os.startfile('Docs\docu.html')
        
    def open_docs(self):
        os.startfile('Docs\how.html')
        
    def open_ver(self):
        os.startfile('Docs\change.txt')
        
    def def_list_close(self):
        self.def_list.close()
    
    def table_close(self, option_tbl):
        option_tbl.close()
     
    def def_close(self):
        self.default_win.close()
    
    def clse2(self):
        self.option_win.close()
        
    #error handlers
    def warnmsg(self, warnstr):
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText(warnstr)
        msg.setWindowTitle("Warning")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        msg.exec_()
        
    def infmsg(self, infstr, sit):
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(infstr)
        msg.setInformativeText(sit)
        msg.setWindowTitle(" ")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setWindowIcon(QtGui.QIcon('media\IDSLOGO.png'))
        msg.exec_()
        
    def errmsg(self, infstr):
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(infstr)
        msg.setWindowTitle(" ")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setWindowIcon(QtGui.QIcon('media\idslogo.jpg'))
        msg.exec_()
    #end of error handlers
def run():
    app = QtGui.QApplication(sys.argv)
    GUI = Window()
    sys.exit(app.exec_())
    
run()