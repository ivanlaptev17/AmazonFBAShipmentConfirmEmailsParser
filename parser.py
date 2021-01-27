import csv
import os
import shutil
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter import messagebox

root = tk.Tk()
root.title('Parser')

class ParsingError(Exception):
    """Basic exception in parsing"""
    pass

class Mail:
    """Class to parse emails easily"""
    def __init__(self, date, sender_info, subject, content):
        self.date = date
        self.sender_info = sender_info
        self.subject = subject
        self.content = content.replace('\r\n', '\n')
    def find(self, ):
        """Returns a dict with values that were found. Others are set as None"""
        found = {'Date': self.date, 'City': None, 'State': None, 'Quantity': None, 'Item': None}
        
        # Finds City and State
        try:
            string = 'to one of your customers:'
            index = self.content.find(string)
            if index != -1:
                to_work = ''
                index += len(string)
                while not self.content[index].lower() in 'abcdefghijklmnopqrstuvwxyz':
                    index += 1
                while self.content[index] != '\n':
                    to_work += self.content[index]
                    index += 1
                found['City'] = to_work.split(',')[0].title()
                found['State'] = to_work.split(',')[1].split()[0].title()
        except Exception:
            raise ParsingError('Unable to find information about State or City')
        
        # Finds Quantity and Item
        try:
            string = 'Item'
            index = self.content.find(string)
            if index != -1:
                to_work = ''
                index += len(string)
                while not self.content[index].isdigit():
                    index += 1
                while self.content[index] != '\n':
                    to_work += self.content[index]
                    index += 1
                found['Quantity'] = int(to_work.split()[0])
                found['Item'] = ' '.join(to_work.split()[1:])
        except Exception:
            raise ParsingError('Unable to find information about State or City')
        
        return found
    staticmethod
    def get_from_dict(d, i):
        """Composes a Mail object from index of mail and dictionary"""
        return Mail(d[date][i], d[sender_details][i], d[subject][i], d[body_contents][i])
    def __str__(self, ):
        """Returns mail (text)"""
        return 'Date: {}\nFrom: {}\nSubject: {}\nBody: \n\n{}'.format(self.date, self.sender_info, self.subject, self.content)

def get_content(file_path):
    """Gets information from csv file and returns it with 'dict' class"""
    global date, sender_details, subject, body_contents
    
    with open(file_path, newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',', quotechar='"')
        lines = list(reader)
        content = {}
        date = lines[0][0]
        sender_details = lines[0][1]
        subject = lines[0][2]
        body_contents = lines[0][3]
    
        for elem in lines[0]:
            content[elem] = []
            
        for line in lines[1:]:
            for i in range(len(content.keys())):
                content[[date, sender_details, subject, body_contents][i]].append(line[i])
    
    return content

def main():
    """Does the parsing and writes data to another csv file"""
    global date, sender_details, subject, body_contents, root, start_button, select_file_button
    
    answer = messagebox.askokcancel('Start Parsing', 'To start parsing you need to choose an output folder')
    if answer != True:
        return None
    file = fd.askdirectory(initialdir='/'.join(file_path.split('/')[-1]))
    if file == '':
        return None
    date = ''
    sender_details = ''
    subject = ''
    body_contents = ''
    content = get_content(file_path)
    
    try:
        os.mkdir(file)
    except FileExistsError:
        pass
    
    with open(file + '/' + 'parsed.csv', 'w', newline='') as csvfile:
        select_file_button['state'] = 'disabled'
        start_button['state'] = 'disabled'
        pb = ttk.Progressbar(length=300)
        lbl_cant_parse = ttk.Label(text='')
        text = ttk.Label(text='Progress: 0%')
        pb.pack()
        text.pack()
        lbl_cant_parse.pack()
        writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        writer.writerow(['Date', 'Time', 'City', 'State', 'Quantity', 'Item'])
        cant_parse = []
        
        for i in range(len(content[date])):
            try:
                mail = Mail.get_from_dict(content, i)
                find = mail.find()
                row = find['Date'].split() + [find['City'], find['State'], find['Quantity'], find['Item']]
                writer.writerow(row)
            except ParsingError as error:
                cant_parse.append([i, str(error)])
                lbl_cant_parse.config(text='Can\'t parse {} letters'.format(len(cant_parse)))
            
            # Updating widgets
            pb['value'] = round(100*i/len(content[date]), 2) - 0.01
            text['text'] = 'Progress: {}%'.format(round(100*(i+1)/len(content[date])))
            
            # Updating root
            root.update()
    
    with open(file + '/Not parsed.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        writer.writerow(['Date', 'Time', 'Error at parsing'])
        for i in cant_parse:
            mail = Mail.get_from_dict(content, i[0])
            writer.writerow(mail.date.split() + [str(i[1])])
    
    pb.destroy()
    text.destroy()
    root.destroy()

def select_file():
    """Selects the file to parse"""
    global file_path, start_button
 
    file_path = fd.askopenfile(defaultextension=('.csv', 'csv files'))
    if file_path == None:
        return None
    file_path = file_path.name
    start_button['state'] = 'normal'

# Creating widgets
select_file_button = ttk.Button(text='Select file', command=select_file)
start_button = ttk.Button(text='Start', command=main)
start_button['state'] = 'disabled'

select_file_button.pack()
start_button.pack()

root.resizable(width=False, height=False)
root.mainloop()