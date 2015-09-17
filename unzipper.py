"""Unzipper

Recursively unzips nested zip files

Usage:
    unzipper [<sourcedir> <destdir>]
    unzipper (-h | --help)

Options:
    -h --help               Show this screen
"""
import os,zipfile,sys,shutil,email,mimetypes,olefile,StringIO
import Tkinter,tkFileDialog,tkMessageBox,ttk

def main (rootdir):
    unzip(rootdir)
    print("FINISHED!")
def unzip (rootdir):
    while (zip_found(rootdir)):
        process_zips(rootdir)
def zip_found(rootdir):
    return_value = False
    for subdirs,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()==".ZIP" or file[-4:].upper()==".MSG":
                return_value = True
    return return_value
def process_zips(rootdir):
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()==".ZIP":
                process_zip(subdir,file)
            elif file[-4:].upper()==".MSG":
                process_msg(subdir,file)
def process_zip(subdir,file):
    os.mkdir(os.path.join(subdir,file)+".dir")
    zip = zipfile.ZipFile(os.path.join(subdir,file))
    zip.extractall(os.path.join(subdir,file)+".dir")
    zip.close()
    os.remove(os.path.join(subdir,file))
def process_mime_msg(subdir,file):
    os.mkdir(os.path.join(subdir,file)+".dir")
    fp = open(os.path.join(subdir,file))
    msg = email.message_from_file(fp)
    fp.close()
    counter = 1
    for part in msg.walk():
        counter = process_mime_msg_section(part,subdir,file,counter)
    os.remove(os.path.join(subdir,file))
def process_mime_msg_section(part,subdir,file,counter):
    if part.get_content_maintype() == 'multipart':
        return counter
    filename = part.get_filename()
    if filename==None:
        filename = generate_mime_msg_section_filename(part,counter)
    counter += 1
    fp = open(os.path.join(os.path.join(subdir,file)+".dir",filename),'wb')
    fp.write(part.get_payload(decode=True))
    fp.close()
    return counter
def generate_mime_msg_section_filename(part,counter):
    ext = mimetypes.guess_extension(part.get_content_type())
    if not ext:
        ext = ".bin"
    filename = "part-{0:03d}{1}".format(counter,ext)
    return filename
def process_msg(subdir,file):
    if olefile.isOleFile(os.path.join(subdir,file))==False:
        process_mime_msg(subdir,file)
        return
    os.mkdir(os.path.join(subdir,file)+".dir")
    ole = olefile.OleFileIO(os.path.join(subdir,file))
    attach_list = get_msg_attach_list(ole)
    extract_msg_files(attach_list,ole,subdir,file)
    extract_msg_message(ole,subdir,file)
    ole.close()
    os.remove(os.path.join(subdir,file))
def get_msg_attach_list(ole):
    attach_list = []
    for i in ole.listdir():
        if i[0][:8]=="__attach":
            if attach_list.count(i[0])==0:
                attach_list.append(i[0])
    return attach_list
def extract_msg_files(attach_list,ole,subdir,file):
    for i in attach_list:
        filename = clean_string(get_msg_attachment_filename(i,ole))
        write_msg_attachment(i,ole,subdir,file,filename)
def get_msg_attachment_filename(index,ole):
    filename = get_msg_attachment_filename_primary(index,ole)
    if filename == None:
        filename = get_msg_attachment_filename_fallback(index,ole)
    if filename == None:
        filename = "ATTACHMENT 1"
    return filename
def get_msg_attachment_filename_primary(index,ole):
    filename = None
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:16]=="__substg1.0_3707":
                name_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                filename = name_stream.read()
    return filename
def get_msg_attachment_filename_fallback(index,ole):
    filename = None
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:16]=="__substg1.0_3704":
                name_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                filename = name_stream.read()
    return filename
def write_msg_attachment(index,ole,subdir,file,filename):
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:20]=="__substg1.0_37010102":
                file_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                file_data = file_stream.read()
                try:
                    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir",filename),"w+b")
                    fp.write(file_data)
                    fp.close()
                except:
                    pass
def extract_msg_message(ole,subdir,file):
    msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body = extract_msg_message_data(ole)    
    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir","00 {0}.txt".format(file)),"w")
    try:
        fp.write("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\nHeader: {4}\n".format(msg_from,msg_to,msg_cc,msg_subject,msg_header).decode('utf-8'))
    except:
        fp.write("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\nHeader: {4}\n".format(msg_from,msg_to,msg_cc,msg_subject,msg_header).decode('ISO-8859-1'))
    fp.write(unicode("---------------\n\n"))
    try:
        fp.write(msg_body.decode('utf-8'))
    except:
        fp.write(msg_body.decode('ISO-8859-1'))
    fp.close()
def extract_msg_message_data(ole):
    msg_from=""
    msg_to=""
    msg_cc=""
    msg_subject=""
    msg_header=""
    msg_body=""
    for i in ole.listdir():
        if i[0][:16]=="__substg1.0_0C1A":
            msg_from = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0E04":
            msg_to = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0E03":
            msg_cc = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0037":
            msg_subject = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_007D":
            msg_header = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_1000":
            msg_body = extract_msg_stream_text(i,ole)
    return (msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body)
def extract_msg_stream_text(index,ole):
    stream = ole.openstream(index[0])
    text = stream.read()
    text = clean_string(text)
    return text
def clean_string(input):
    output = ""
    save_flag = False
    for letter in input:
        if save_flag == False:
            save_flag = True
        else:
            save_flag = False
        if save_flag == True:
            output = output + letter    
    return output
def launch_main(sourcedir,destdir):
    if sourcedir==None or destdir==None:
        print("One or more missing arguments. Exiting.")
        sys.exit()
    sourcedir = os.path.abspath(sourcedir)
    destdir = os.path.abspath(destdir)
    os.rmdir(destdir)
    shutil.copytree(sourcedir,destdir)
    main(destdir)
def launch_gui():
    root = Tkinter.Tk()
    root.title("Unzipper")
    global app
    app = Application(master=root)
    app.mainloop()
    root.destroy()
class Application(Tkinter.Frame):
    def __init__(self, master = None):
        Tkinter.Frame.__init__(self,master)
        self.pack()
        self.source_directory = ""
        self.dest_directory = ""
        self.page_setup_settings = None
        self.create_widgets()
    def create_widgets(self):
        self.create_exit()
        self.create_start()
        self.create_choose_source()
        self.create_source_text()
        self.create_choose_dest()
        self.create_dest_text()
    def create_exit(self):
        self.exit = Tkinter.Button(self)
        self.exit["text"] = "Exit"
        self.exit["command"] = self.quit
        self.exit.grid(row=5,column=1)
    def create_start(self):
        self.start_button = Tkinter.Button(self)
        self.start_button["text"] = "Start"
        self.start_button["command"] = self.start
        self.start_button.grid(row=5,column=0)
    def create_choose_source(self):
        self.choose_source = Tkinter.Button(self)
        self.choose_source["text"] = "Source Directory:"
        self.choose_source["command"] = self.source_directory_select
        self.choose_source.grid(row=0,column=0)
    def create_choose_dest(self):
        self.choose_dest = Tkinter.Button(self)
        self.choose_dest["text"] = "Destination Directory:"
        self.choose_dest["command"] = self.dest_directory_select
        self.choose_dest.grid(row=1,column=0)
    def create_source_text(self):
        self.chosen_source = Tkinter.Label(self)
        self.chosen_source["text"] = self.source_directory
        self.chosen_source.grid(row=0,column=1)
    def create_dest_text(self):
        self.chosen_dest = Tkinter.Label(self)
        self.chosen_dest["text"] = self.dest_directory
        self.chosen_dest.grid(row=1,column=1)
    def source_directory_select(self):
        self.source_directory = tkFileDialog.askdirectory(initialdir = os.getcwd(), title = "Choose Source Directory", mustexist=True)
        self.chosen_source["text"] = self.source_directory
    def dest_directory_select(self):
        self.dest_directory = tkFileDialog.askdirectory(initialdir = os.getcwd(), title = "Choose Destination Directory", mustexist=True)
        self.chosen_dest["text"] = self.dest_directory
    def start(self):
        if self.source_directory==None or self.dest_directory==None or self.source_directory=="" or self.dest_directory=="":
            tkMessageBox.showerror("Error","Missing fields - cannot launch.")
        elif self.source_directory=="C:\\" or self.source_directory=="C:/" or self.source_directory=="C:" or self.dest_directory=="C:\\" or self.dest_directory=="C:/" or self.dest_directory=="C:":
            tkMessageBox.showerror("Error","Will not launch using the root directory.")
        elif os.listdir(self.dest_directory) != []:
            tkMessageBox.showerror("Error","Destination directory must be empty.")
        elif self.source_directory==self.dest_directory:
            tkMessageBox.showerror("Error","Source and destination directories cannot be the same.")
        else:
            launch_main(self.source_directory,self.dest_directory)
    
if __name__ == "__main__":
    from docopt import docopt
    arguments = docopt(__doc__, version='RRR 0.1')
    sourcedir = arguments["<sourcedir>"]
    if sourcedir == "C:\\" or sourcedir == "C:" or sourcedir=="C:/": #this would be bad
        print("Will not run on root directory. Exiting.")
        sys.exit()
    destdir = arguments["<destdir>"]
    if destdir == "C:\\" or destdir == "C:" or destdir=="C:/": #this would be bad
        print("Will not run on root directory. Exiting.")
        sys.exit()
    if destdir != None and os.listdir(destdir) != []:
        print("Destination directory must be empty. Exiting.")
        sys.exit()
    if sourcedir==None or destdir==None:
        launch_gui()
        sys.exit()
    if sourcedir==destdir:
        print("Source and destination directories are the same. Exiting.")
        sys.exit()
    launch_main(sourcedir,destdir)
