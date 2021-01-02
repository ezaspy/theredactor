#!/usr/bin/env python3 -tt
import argparse, os, re, sys, shutil, string, random, subprocess, hashlib, datetime
from zipfile import ZipFile
from tkinter import Tk, PhotoImage, Label, Button, Entry, filedialog, messagebox

parser = argparse.ArgumentParser()
parser.add_argument("directory", nargs=1, help="Source directory where the document(s) are located.")

args = parser.parse_args()
d = args.directory

s_quote = ["Judgement Day is inevitable.", "It is time.", "You must live.", "Cyberdyne Systems, Model 101.", "Thirty-five years from now, you reprogrammed me to be your protector here, in this time.", "The T-1000 will definitely try to reacquire you.", "I swear I will not kill anyone.", "Skynet presets the switch to read-only.", "It launches its missiles against the targets in Russia.", "Negative. She's not a mission priority.", "Trust me.", "I need your clothes, your boots and your motorcycle.", "Come with me if you want to live!", "I have to stay functional until my mission is complete.", "My CPU is a neural-net processor; a learning computer.", "My mission is to protect you."]
e_quote = ["Judgement Day is inevitable.", "The more contact I have with humans, the more I learn.", "I need a vacation.", "It's in your nature to destroy yourselves.", "Typically, the subject being copied is terminated.", "I'll be back.", "Hasta la vista, baby.", "No problemo.", "Terminated.", "He'll live.", "I'll take care of the police."]
xd, hashlist, common_words, r, a, i, contents, password = ([] for i in range(8))
lines, proc = ({} for i in range(2))
now = datetime.datetime.now()
dictionary = os.path.join("media", "dictionary")
hashes = os.path.join("media", "hashes")
log = os.path.join("media", "log")
re_md5 = r"[\w]{32}"
re_acct = r"(?:[\d\ ]{1,4}\ |K\ ?ZT\d?|\d[A-Za-z\ ])(?:[\d\ ]{2,5})(?:[A-Za-z]{1,3}|\d{1})?(?:[A-Za-z]|\d{1})?\ ?(?:\d{1,2})?\ ?(?:[A-Za-z]|\d{1,4})?\ ?(?:[A-Za-z]{2})?\ ?(?:[A-Za-z]{1,4}|\d{1,4})?\ ?(?:\d{1,2})?(?:[A-Za-z]{1}|\d{2})?"
re_bban = r"(?:[A-Za-z]{1,4}|\d{2})\ ?(?:[\d\ ]{4}|K\ ?ZT\d?|\d[A-Za-z])\ ?(?:[A-Za-z]{2}|\d{1,5})\ ?(?:\d{1,3})(?:[A-Za-z]|\d)\ ?(?:[A-Za-z]|\d{1,4})(?:[A-Za-z]{1,2}|\d{1,3})\ ?(?:[A-Za-z]{1,3}|\d{1,2})\ ?(?:[A-Za-z]{1,2}|\d{1,2})?\ ?(?:[A-Za-z]{4}|\d{2,4})?\ ?(?:[A-Za-z]{1}|\d{1,4})?\ ?(?:[A-Za-z]{1,2}|\d{1})?(?:\ \d{1,2}[A-Za-z])?"
re_cn = r"[\ |\,]{0,1}[\d]{13,18}[\ |\,]{0,1}"
re_ea = r"[\w\d\-\.]{1,63}\[?\@\]?[\w\-]+\[?\.\]?[\w-]+(?:[\w\-\.]+)?(?:[\w\-\.]+)?"
re_hn = r"[A-Za-z]{7}\d{6}"
re_iban = r"(?:AD|AE|AL|AT|AZ|BA|BE|BG|BH|BR|CH|CR|CY|CZ|DE|DK|DO|EE|ES|FI|FO|FR|GB|GE|GI|GL|GR|GT|HR|HU|IE|IL|IS|IT|JO|KW|KZ|LB|LC|LI|LT|LU|LV|MC|MD|ME|MK|MR|MT|MU|NL|NO|PK|PL|PS|PT|QA|RO|RS|SA|SE|SI|SK|SM|ST|TL|TN|TR|VG|XK)\d{2}\ ?(?:[A-Za-z]{1,4}|\d{2})\ ?(?:[\d\ ]{4}|K\ ?ZT\d?|\d[A-Za-z])\ ?(?:[A-Za-z]{2}|\d{1,5})\ ?(?:\d{1,3})(?:[A-Za-z]|\d)\ ?(?:[A-Za-z]|\d{1,4})(?:[A-Za-z]{1,2}|\d{1,3})\ ?(?:[A-Za-z]{1,3}|\d{1,2})\ ?(?:[A-Za-z]{1,2}|\d{1,2})?\ ?(?:[A-Za-z]{4}|\d{2,4})?\ ?(?:[A-Za-z]{1}|\d{1,4})?\ ?(?:[A-Za-z]{1,2}|\d{1})?(?:\ \d{1,2}[A-Za-z])?"
re_imei = r"[\ |\,]{0,1}[\d]{15}[\ |\,]{0,1}"
re_ip4 = r"(?:2|1)?(?:\d)?\d\.(?:2|1)?(?:\d)?\d\.(?:2|1)?(?:\d)?\d+\.(?:2|1)?(?:\d)?\d"
re_mac = r"[A-Fa-f\d]{2}:[A-Fa-f\d]{2}:[A-Fa-f\d]{2}:[A-Fa-f\d]{2}:[A-Fa-f\d]{2}:[A-Fa-f\d]{2}"
re_un = r"(?:[A-Za-z]{3}?)?(?:\\|\/)?(?:[A-Za-z\_]{4})?(?:[A-Za-z]{5,9})(?:[\_\-])?(?:[A-Za-z]{1,2})?"
re_phone = r"[\+]?[\d\ \-\_\(\)]{6,20}"
re_swift = r"[\ |\,]{0,1}[A-Za-z]{6}\d{2}(?:[A-Za-z\d]{3})?[\ |\,]{0,1}"
re_pm = r"((?i)(?:strictly\s)?confidential|(?i)internal)"

def doWindow(stage, *rest):
    def doEntry():
        password.append(text.get())
        doClose()
    def doAbout():
        about = messagebox.showinfo("The Redactor - About", "The Redactor can read PII from: \nCSV (.csv)\nText (.txt)\nOffice Documents (.xlsx, .docx)\n\nOnly the following forms of PII can be extracted:\nBank Account Number\nBBAN (Basic Bank Account Number)\nCredit/Debit Card Number\nEmail Address\nHostname\nIBAN (International Bank Account Number)\nIMEI\nIPv4 Address\nMAC Address\nUsername\nInternational Phone Number\nSWIFT Code\n\nThe Redactor does not identify all PII, including but not limited too:\nFull name\nAddress\nNational identification number\nVehicle registration plate\nPassport numbers\nCryptocurrency identifiers\nCookies Identifiers")
        if about == "ok":
            pass
        else:
            pass
    def doClose():
        window.destroy()
    def doExit():
        window.destroy()
        sys.exit()
    window = Tk()
    window.title("The Redactor")
    window.geometry("820x360+100+60")
    window.configure(background="black")
    Label(window, text="The Redactor", bg="black", fg="white", font=("Verdana bold", 40)).place(x=260, y=20)
    Label(window, text=random.choice(s_quote), bg="black", fg="green", font="Calibri 24").place(x=100, y=110)
    Label(window, text="In accordance with GDPR, this program searches files containing Personally Identifiable Information from said file(s) and redacts that data if requested.", bg="black", fg="white", font="Calibri 12").place(x=10, y=160)
    Label(window, text="Please note; this tool is NOT foolproof and should not be used as the sole means for checking if a file has had all sensitive information removed.", bg="black", fg="red", font="Calibri 12").place(x=10, y=180)
    Button(window, text="About", fg="black", font="Calibri 20", command=doAbout).place(x=100, y=20)
    Button(window, text="Exit", fg="black", font="Calibri 20", command=doExit).place(x=20, y=20)
    if stage == "welcome":
        Button(window, text="Redact", fg="black", font="Calibri 40", command=doClose, width=18).place(x=230, y=250)
    elif stage == "repeat":
        restpath, restname = str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][0]), str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][1])
        if restname.endswith("bugs") or restname.endswith("UserGuide.docx") or restname.endswith("redactor.exe") or restname.endswith("redactor.py") or restname.endswith("hashes") or restname.endswith("dictionary") or restname.endswith("log") or restname.startswith("~"):
            doClose()
        else:
            Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=90, y=250)
            Label(window, text="File: '%s'" % (restname), bg="black", fg="white", font="Calibri 16").place(x=90, y=280)
            result = messagebox.askyesno("The Redactor", "The Redactor has reviewed this file before.\nDo you wish to process the file again?")
            if result == True:
                result = "y"
            else:
                result = "n"
            r.append(result)
            doClose()
    elif stage == "enc":
        restpath, restname = str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][0]), str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][1])
        attempt = rest[1]
        if attempt == 0:
            Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
            Label(window, text="File: '%s' is encrypted." % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
        else:
            Label(window, text="Password is incorrect, please try again for:", bg="black", fg="white", font="Calibri 16").place(x=110, y=250)
            Label(window, text="'%s'" % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=280)
        Label(window, text="Please enter the password:", bg="black", fg="red", font="Calibri 16").place(x=110, y=330)
        text = Entry(window, width=40)
        text.place(x=380, y=340)
        Button(window, text="Decrypt", fg="black", font="Calibri 20", command=doEntry).place(x=640, y=320)
    elif stage == "ans":
        pm_label = rest[1]
        if rest[2] == "Protective Marking":
            restpath, restname = str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][0]), str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][1])
            if pm_label == "Internal":
                pm_label = "For internal use only"
            else:
                pass
            Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
            Label(window, text="File: '%s' is potentially protectively marked as:" % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
            Label(window, text="'%s'" % (pm_label), bg="black", fg="white", font="Calibri 16").place(x=110, y=300)
            result = messagebox.askyesno("The Redactor", "Do you wish to continue?")
        else:
            content, field, linecell = rest[3], rest[2].strip(" ,"), "line with"
            if "document.xml" in content:
                re_line = r"\<[^\<]+"+re.escape(rest[1])+r"[^\>]+\>"
                lines = re.findall(re_line, rest[0])
                if lines:
                    line = lines[0][5:-6]
                    line = line.replace('"', "'").split("xml:space='preserve'>")
                    if len(line) == 2:
                        line = line[1]
                    else:
                        line = line[0]
                    filename = rest[3].rsplit("/word")
                    filename = filename[0]+".docx"
                else:
                    line = "|0|"
            elif "sharedStrings.xml" in content:
                linecell = "cell with"
                re_line = r"\>\<t\>"+re.escape(rest[1])+r"\<\/t\>\<"
                lines = re.findall(re_line, rest[0])
                if lines:
                    line = lines[0][4:-5]
                    line = line.replace('"', "'").split("xml:space='preserve'>")
                    if len(line) == 2:
                        line = line[1]
                    else:
                        line = line[0]
                    filename = rest[3].rsplit("/xl")
                    filename = filename[0]+".xlsx"
                else:
                    line = "|0|"
            else:
                line = rest[0]
                filename = rest[3]
            if line != "|0|":
                restpath = rest[3].split("/")
                restpath = restpath[0]
                restname = filename.split("/")[1]
                if content == "IBAN (International Bank Account Number)":
                    Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
                    Label(window, text="File: '%s'" % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
                    Label(window, text="For %s: '%s'" % (linecell, line), bg="black", fg="white", font="Calibri 16").place(x=110, y=300)
                    Label(window, text="'%s' has been detected as an '%s' from '%s'" % (rest[1], rest[4], rest[2]), bg="black", fg="white", font="Calibri 16").place(x=110, y=300)
                    result = messagebox.askyesno("The Redactor", "Do you wish to redact this data?")
                elif content == "Card Number":
                    Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
                    Label(window, text="File: '%s'" % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
                    Label(window, text="For %s: '%s'" % (linecell, line), bg="black", fg="white", font="Calibri 16").place(x=110, y=300)
                    Label(window, text="'%s' has been detected as a '%s'" % (rest[1], rest[3]), bg="black", fg="white", font="Calibri 16").place(x=110, y=330)
                    result = messagebox.askyesno("The Redactor", "Do you wish to redact this data?")
                else:
                    if content == "an Account Number" or content =="an IMEI number":
                        field = field.strip(" ,<>")
                    else:
                        pass
                    Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
                    Label(window, text="File: '%s'" % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
                    Label(window, text="For %s: '%s'" % (linecell, line), bg="black", fg="white", font="Calibri 16").place(x=110, y=300)
                    Label(window, text="'%s' has been detected as '%s'" % (rest[1], rest[2]), bg="black", fg="white", font="Calibri 16").place(x=110, y=330)
                    result = messagebox.askyesno("The Redactor", "Do you wish to redact this data?")
            else:
                result = False
        if result == True:
            result_all = messagebox.askyesno("The Redactor", "Do you wish to redact all instances of this data?")
            if result_all == True:
                ans = "a"
                a.append(rest[1])
            else:
                ans = "y"
        else:
            result_all = messagebox.askyesno("The Redactor", "Do you wish to ignore all further instances of this data?")
            if result_all == True:
                ans = "i"
                i.append(rest[1])
            else:
                ans = "n"
        r.append(ans)
        doClose()
    elif stage == "invalid":
        restpath, restname = str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][0]), str(re.findall(r"([\S\s]+\/)([^\/]+$)", rest[0])[0][1])
        if rest[0].endswith("bugs") or rest[0].endswith("UserGuide.docx") or rest[0].endswith("redactor.exe") or rest[0].endswith("redactor.py") or rest[0].endswith("hashes") or rest[0].endswith("dictionary") or rest[0].endswith("log"):
            doClose()
        elif restname.startswith("~"):
                Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
                Label(window, text="File: '%s' is currently open." % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
                result = messagebox.showinfo("File is Open", "%s\nis currently open and must be closed before The Redactor can run.\nPlease try again." % (rest[0]))
                if result == "ok":
                    doExit()
                else:
                    pass
        elif rest[0].endswith(".pdf") or rest[0].endswith(".jpg") or rest[0].endswith(".jpeg") or rest[0].endswith(".png") or rest[0].endswith(".gif"):
            if rest[0].endswith(".pdf"):
                Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
                Label(window, text="File: '%s' has been identified as a PDF Document." % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
                result = messagebox.showinfo("PDF Found", "%s\nMay need further inspection" % (rest[0]))
            elif rest[0].endswith(".jpg") or rest[0].endswith(".jpeg") or rest[0].endswith(".png") or rest[0].endswith(".gif"):
                Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
                Label(window, text="File: '%s' has been identified as an image file." % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
                result = messagebox.showinfo("Image Found", "%s\nMay need further inspection" % (rest[0]))
            root = rest[0].rsplit("/")
            dup = "/OUT__"+root[-1]
            root.pop()
            dup = '/'.join(root)+dup
            shutil.copy2(rest[0], dup)
            if result == "ok":
                doClose()
            else:
                pass
        else:
            Label(window, text="Path: '%s'" % (restpath), bg="black", fg="white", font="Calibri 16").place(x=110, y=240)
            Label(window, text="File: '%s' is not a valid file type." % (restname), bg="black", fg="white", font="Calibri 16").place(x=110, y=270)
            result = messagebox.showinfo("Invalid File Type", "PII can only be extracted from the following files:\nCSV (.csv)\nText (.txt)\nOffice Documents (.xlsx, .docx)")
            if result == "ok":
                doClose()
            else:
                pass
    elif stage == "review":
        common_wordsset = set(rest[0])
        Label(window, text="The following fields have been identified as potential common words", bg="black", fg="white", font="Calibri 16").place(x=110, y=250)
        Label(window, text="Please review...", bg="black", fg="white", font="Calibri 16").place(x=110, y=280)
        for each_word in common_wordsset:
            result = messagebox.askyesno("The Redactor", "Is '%s' a too common word?" % each_word.strip())
            if result == True:
                with open(dictionary, "a") as w:
                    w.write(each_word.strip()+"\n")
            else:
                pass
        doClose()
    elif stage == "exit":
        Label(window, text="A directory was not selected or the program was closed. Please try again.", bg="black", fg="white", font="Calibri 16").place(x=110, y=250)
        Label(window, text=random.choice(e_quote), bg="black", fg="green", font="Calibri 16").place(x=110, y=290)
    elif stage == "finished":
        Label(window, text="The Redactor has completed successfully. Please exit the program.", bg="black", fg="white", font="Calibri 16").place(x=110, y=250)
        Button(window, text="Exit", fg="black", font="Calibri 28", command=doClose, width=18).place(x=320, y=300)
    else:
        Label(window, text="An error occured. Please close this window.", bg="black", fg="white", font="Calibri 16").place(x=110, y=250)
        Button(window, text="Exit", fg="black", width=6, font="Calibri 20", command=doClose).place(x=60, y=30)
    window.mainloop()
def doQstn(stage, line, field, content, x, seen):
    def IBAN(field):
        global country
        field = field.upper()
        if field[:2] == "AD":
            country = "Andorra"
        elif field[:2] == "AE":
            country = "United Arab Emirates"
        elif field[:2] == "AL":
            country = "Albania"
        elif field[:2] == "AT":
            country = "Austria"
        elif field[:2] == "AZ":
            country = "Azerbaijan"
        elif field[:2] == "BA":
            country = "Bosnia and Herzegovina"
        elif field[:2] == "BE":
            country = "Belgium"
        elif field[:2] == "BG":
            country = "Bulgaria"
        elif field[:2] == "BH":
            country = "Bahrain"
        elif field[:2] == "BR":
            country = "Brazil"
        elif field[:2] == "BY":
            country = "Belarus"
        elif field[:2] == "CH":
            country = "Switzerland"
        elif field[:2] == "CR":
            country = "Costa Rica"
        elif field[:2] == "CY":
            country = "Cyprus"
        elif field[:2] == "CZ":
            country = "Czech Republic"
        elif field[:2] == "DE":
            country = "Germany"
        elif field[:2] == "DK":
            country = "Denmark"
        elif field[:2] == "DO":
            country = "Dominican Republic"
        elif field[:2] == "EE":
            country = "Estonia"
        elif field[:2] == "ES":
            country = "Spain"
        elif field[:2] == "FI":
            country = "Finland"
        elif field[:2] == "FO":
            country = "Faroe Islands"
        elif field[:2] == "FR":
            country = "France, French OST or Madagascar"
        elif field[:2] == "GB":
            country = "United Kingdom"
        elif field[:2] == "GE":
            country = "Georgia"
        elif field[:2] == "GI":
            country = "Gibraltar"
        elif field[:2] == "GL":
            country = "Greenland"
        elif field[:2] == "GR":
            country = "Greece"
        elif field[:2] == "GT":
            country = "Guatemala"
        elif field[:2] == "HR":
            country = "Croatia"
        elif field[:2] == "HU":
            country = "Hungary"
        elif field[:2] == "IE":
            country = "Ireland"
        elif field[:2] == "IL":
            country = "Israel"
        elif field[:2] == "IS":
            country = "Iceland"
        elif field[:2] == "II":
            country = "Italy"
        elif field[:2] == "JO":
            country = "Jordan"
        elif field[:2] == "KW":
            country = "Kuwait"
        elif field[:2] == "KZ":
            country = "Kazakhstan"
        elif field[:2] == "LB":
            country = "Lebanon"
        elif field[:2] == "LB":
            country = "Saint Lucia"
        elif field[:2] == "LI":
            country = "Liechenstein"
        elif field[:2] == "LT":
            country = "Lithuania"
        elif field[:2] == "LU":
            country = "Luxembourg"
        elif field[:2] == "LV":
            country = "Latvia"
        elif field[:2] == "MC":
            country = "Monaco"
        elif field[:2] == "MD":
            country = "Moldova"
        elif field[:2] == "ME":
            country = "Montenegro"
        elif field[:2] == "MK":
            country = "Macedonia"
        elif field[:2] == "MR":
            country = "Mauritania"
        elif field[:2] == "MT":
            country = "Malta"
        elif field[:2] == "MU":
            country = "Mauritius"
        elif field[:2] == "NL":
            country = "The Netherlands"
        elif field[:2] == "NO":
            country = "Norway"
        elif field[:2] == "PK":
            country = "Pakistan"
        elif field[:2] == "PL":
            country = "Poland"
        elif field[:2] == "PS":
            country = "Palestine"
        elif field[:2] == "PT":
            country = "Portugal"
        elif field[:2] == "QA":
            country = "Qatar"
        elif field[:2] == "RO":
            country = "Romania"
        elif field[:2] == "RS":
            country = "Serbia"
        elif field[:2] == "SA":
            country = "Suadi Arabia"
        elif field[:2] == "SE":
            country = "Sweden"
        elif field[:2] == "SI":
            country = "Slovenia"
        elif field[:2] == "SK":
            country = "Slovakia"
        elif field[:2] == "SM":
            country = "San Marino"
        elif field[:2] == "ST":
            country = "Sao Tome and Principe"
        elif field[:2] == "TL":
            country = "Timor Leste"
        elif field[:2] == "TN":
            country = "Tunisia"
        elif field[:2] == "TR":
            country = "Turkey"
        elif field[:2] == "UA":
            country = "Ukraine"
        elif field[:2] == "VG":
            country = "British Virgin Islands"
        elif field[:2] == "XK":
            country = "Kosovo"
        else:
            country = "an unknown origin"
    def luhn(field):
        def doIIN(luhnout, field):
            global iin
            if luhnout == True:
                if field[:1] == "1":
                    iin = "UATP"
                elif field[:1] == "2":
                    if field[:4] == "2014" or field[:4] == "2149":
                        iin = "DINERS CLUB ENROUTE"
                    elif 2200 <= int(field[:4]) <= 2204:
                        iin = "MIR"
                    elif 222100 <= int(field[:6]) <= 272099:
                        iin = "MASTERCARD"
                    else:
                        iin = "UNKNOWN CARD ISSUER"
                elif field[:1] == "3":
                    if field[:2] == "34" or field[:2] == "37":
                        iin = "AMEX"
                    elif field[:2] == "36" or field[:2] == "38" or field[:2] == "39" or field[:4] == "3095":
                        iin = "DINERS CLUB INTERNATIONAL"
                    elif 300 <= int(field[:3]) <= 305:
                        iin = "DINERS CLUB INTERNATIONAL"
                    elif 3528 <= int(field[:4]) <= 3589:
                        iin = "JCB"
                    else:
                        iin = "UNKNOWN CARD ISSUER"
                elif field[:1] == "4":
                    if field[:4] == "4903" or field[:4] == "4905" or field[:4] == "4911" or field[:4] == "4936":
                        iin = "SWITCH"
                    elif field[:4] == "4571":
                        iin = "DANKORT"
                    else:
                        iin = "VISA"
                elif field[:1] == "5":
                    if field[:4] == "5610":
                        iin = "BANKCARD"
                    elif 560221 <= int(field[:6]) <= 560225:
                        iin = "BANKCARD"
                    elif field[:2] == "54" or field[:2] == "55":
                        iin = "DINERS CLUB (US & CA)"
                    elif field[:2] == "50" or field[:2] == "56" or field[:2] == "57" or field[:2] == "58":
                        iin = "MAESTRO"
                    elif field[:4] == "5019":
                        iin = "DANKORT"
                    elif 51 <= int(field[:2]) <= 55:
                        iin = "MASTERCARD"
                    elif field[:6] == "564182":
                        iin = "SWITCH"
                    elif 506099 <= int(field[:6]) <= 506198:
                        iin = "VERVE"
                    else:
                        iin = "UNKNOWN CARD ISSUER"
                elif field[:1] == "6":
                    if field[:2] == "62":
                        iin = "CHINA UNIONPAY"
                    elif field[:4] == "6011" or field[:2] == "64" or field[:2] == "65":
                        iin = "DISCOVER CARD"
                    elif field[:2] == "60" or field[:4] == "6521":
                        iin = "RUPAY"
                    elif field[:3] == "636":
                        iin = "INTERPAYMENT"
                    elif field[:3] == "637" or field[:3] == "638" or field[:3] == "639":
                        iin = "INSTAPAYMENT"
                    elif field[:4] == "6304" or field[:4] == "6706" or field[:4] == "6771" or field[:4] == "6709":
                        iin = "LASER"
                    elif field[:2] == "67":
                        iin = "MAESTRO"
                    elif field[:4] == "6334" or field[:4] == "6767":
                        iin = "SOLO"
                    elif field[:6] == "633110" or field[:4] == "6333" or field[:4] == "6759":
                        iin = "SWITCH"
                    elif 650002 <= int(field[:6]) <= 650027:
                        iin = "VERVE"
                    else:
                        iin = "UNKNOWN CARD ISSUER"
                elif field[:1] == "9":
                    if 979200 <= int(field[:6]) <= 979289:
                        iin = "TROY"
                    else:
                        iin = "UNKNOWN CARD ISSUER"
                else:
                    iin = "UNKNOWN CARD ISSUER"
            else:
                iin = "UNKNOWN CARD ISSUER"
            return iin
        global luhnout
        r = [int(ch) for ch in str(field)][::-1]
        luhnout = (sum(r[0::2]) + sum(sum(divmod(d*2,10)) for d in r[1::2])) % 10 == 0
        doIIN(luhnout, field)
    global ans
    stage == "ans"
    if content == "Protective Marking":
        field = field.title()
        doWindow(stage, x, field, content)
    else:
        if seen == "no":
            if content == "Card Number":
                field = field.strip(" ,")
                luhn(field)
                if luhnout == True:
                    doWindow(stage, line.strip(), field, content, x, iin)
                else:
                    ans = "n"
            else:
                if content == "IBAN (International Bank Account Number)":
                    field = field.strip(" ,")
                    IBAN(field)
                    doWindow(stage, line.strip(), field, content, x, country)
                elif content == "an Account Number" or content == "a BBAN (Basic Bank Account Number)" or content == "an IMEI number":
                    field = field.strip(" ,")
                    doWindow(stage, line.strip(), field, content, x)
                else:
                    doWindow(stage, line.strip(), field, content, x)
        else:
            pass
    ans = r[-1]
    return ans
def doPrctvMrkg(root, x, filein, filetype):
    for line in filein:
        if re.findall(re_pm, line):
            fields = re.findall(re_pm, line)
            for field in fields:
                if "xxxx" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                    stage = "ans"
                    content = "Protective Marking"
                    seen = "no"
                    doQstn(stage, line, field, content, x, seen)
                    if ans == "n":
                        sys.exit()
                    else:
                        pass
                else:
                    pass
        else:
            pass
def doRedact(x, line, line_no):
    global lines
    global seen
    def doCompare(field):
        global m
        m = ""
        with open(dictionary, "r") as r:
            for f in r:
                f = f.strip("\n")
                if field in f:
                    if len(f) == len(field):
                        m = "match"
                    else:
                        m = "no match"
                else:
                    m = "no match"
        return m
    def doLog(content, field, x, ans, line):
        if ans != "n":
            replace = "yes"
        else:
            replace = "no"
        entry = "\""+now.isoformat()+"\",\""+content+"\",\""+field+"\",\""+x+"\","+replace+"\",\""+line.strip()+"\"\n"
        with open(log, "a") as logentry:
            logentry.write(entry)
    if re.findall(re_acct, line) or re.findall(re_bban, line) or re.findall(re_cn, line) or re.findall(re_ea, line) or re.findall(re_hn, line) or re.findall(re_iban, line) or re.findall(re_imei, line) or re.findall(re_ip4, line) or re.findall(re_mac, line) or re.findall(re_un, line) or re.findall(re_phone, line):
        fields = re.findall(re_iban, line)
        for field in fields:
            doCompare(field)
            string = "xxxx xxxx xxxx xxxx xxxx xxxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "IBAN (International Bank Account Number)"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field.strip(" ,"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_bban, line)
        for field in fields:
            doCompare(field)
            string = "xxxx xxxx xxxx xxxx xxxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "a BBAN (Basic Bank Account Number)"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field.strip(" ,"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_acct, line)
        for field in fields:
            doCompare(field)
            string = "xxxxxxxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx" and field != "0C71FC0" and field != "0D91580" and field != "0C76642" and field!="" and field!=" " and field!="  " and field!="   " and field!="    " and field!="     " and field!="      " and field!="       " and field!="        " and field!="         " and field!="          " and field!="           ":
                stage = "ans"
                content = "an Account Number"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field.strip(" ,"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_cn, line)
        for field in fields:
            doCompare(field)
            string = "xxxx xxxx xxxx xxxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "Card Number"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)

                if ans != "n":
                    line = re.sub(field.strip(" ,"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_ea, line)
        for field in fields:
            doCompare(field)
            string = "x@x.x"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "an Email Address"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field, string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_hn, line)
        for field in fields:
            doCompare(field)
            string = "xxxxxxxxxxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "a Hostname"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field, string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_imei, line)
        for field in fields:
            doCompare(field)
            string = "xxxxxxxxxxxxxxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "an IMEI number"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field.strip(" ,"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_ip4, line)
        for field in fields:
            doCompare(field)
            string = "xxx.xxx.xxx.xxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "an IPv4 Address"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field, string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_mac, line)
        for field in fields:
            doCompare(field)
            string = "xx:xx:xx:xx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "a MAC Address"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field, string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_un, line)
        for field in fields:
            doCompare(field)
            string = "x"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                if m == "no match":
                    with open(dictionary, "r+") as words:
                        counter = 0
                        wordlist = []
                        low_len = subprocess.Popen(["find", "/v", "/c", "", dictionary], stdout=subprocess.PIPE)
                        out = low_len.communicate()
                        out = str(out)
                        out = re.findall(r"\d+", out)
                        out = out[0]
                        out = int(out)
                        while counter < out:
                            for word in words:
                                word = word.strip("\n")
                                counter += 1
                                if field.lower() != word.lower():
                                    wordlist.append("no match")
                                else:
                                    wordlist.append("match")
                        wordset = set(wordlist)
                        wordmatch = len(wordset)
                        if wordmatch == 1:
                            stage = "ans"
                            content = "a Username"
                            seen = "no"
                            if len(a) > 0:
                                for each in a:
                                    if each == field:
                                        seen = "yes"
                                    else:
                                        pass
                            else:
                                pass
                            if len(i) > 0:
                                for each in i:
                                    if each == field:
                                        seen = "ignore"
                                    else:
                                        pass
                            else:
                                pass
                            doQstn(stage, line, field, content, x, seen)
                            if ans == "y" or ans == "a":
                                line = re.sub(field, string, line)
                            else:
                                common_words.append(field+"\n")
                            doLog(content, field, x, ans, line)
                        else:
                            pass
                else:
                    pass
            else:
                pass
        fields = re.findall(re_phone, line)
        for field in fields:
            doCompare(field)
            string = "xx xxxx xxx xxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "a Phone Number"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = line.replace("(", "_").replace(")", "_")
                    field = field.replace("(", "_").replace(")", "_")
                    line = re.sub(field.strip(" +"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        fields = re.findall(re_swift, line)
        for field in fields:
            doCompare(field)
            string = "xxxxxxxx-xxx"
            if "xxxx" not in field and "onfidential" not in field and "nternal" not in field and field != "xx-xx-xxxx" and field != "x@x.x" and field != "xxx.xxx.xxx.xxx" and field != "xx:xx:xx:xx" and field != "+xx xxxx xxx xxx" and field != "x" and field != "xxxxxxxx-xxx":
                stage = "ans"
                content = "a SWIFT Code"
                seen = "no"
                if len(a) > 0:
                    for each in a:
                        if each == field:
                            seen = "yes"
                        else:
                            pass
                else:
                    pass
                if len(i) > 0:
                    for each in i:
                        if each == field:
                            seen = "ignore"
                        else:
                            pass
                else:
                    pass
                doQstn(stage, line, field, content, x, seen)
                if ans != "n":
                    line = re.sub(field.strip(" ,"), string, line)
                else:
                    pass
                doLog(content, field, x, ans, line)
            else:
                pass
        lines[line_no] = line.strip()
    else:
        lines[line_no] = line.strip()
def doWrite(root, x, filein, filetype):
    def makeRedactedFile(filein):
        line_no = 0
        for line in filein:
            doRedact(x, line, line_no)
            line_no += 1
    if filetype != "office":
        makeRedactedFile(filein)
        y = x.rsplit("/", 1)[-1]
        y = root+"/OUT__"+y
        with open(y, "w") as fileout:
            for key in sorted(lines):
                fileout.write(lines[key]+"\n")
    else:
        if "document" in x:
            for line in filein:
                struct = line.rsplit("w:body>")
            filein = []
            filein.append(struct[1])
            makeRedactedFile(filein)
            with open(x, "w") as fileout:
                for key in sorted(lines):
                    fileout.write(struct[0]+"w:body>"+lines[key]+"w:body>"+struct[2])
            y = root+"/OUT__"+x.split(root)[1].split("/")[1]+".docx"
            z = ZipFile(y, "w")
            zippath = x.rsplit("/word")
            for ziproot, _, zipfiles in os.walk(zippath[0]):
                for f in zipfiles:
                    path = ziproot+"/"+f
                    if len(path.split(zippath[0])[1].split("\\")) == 3:
                        name = "/"+path.split(zippath[0])[1].split("\\")[1]+"/"+path.split(zippath[0])[1].split("\\")[2]
                    elif len(path.split(zippath[0])[1].split("\\")) == 2:
                        name = "/"+path.split(zippath[0])[1].split("\\")[1]
                    else:
                        name = path.split(zippath[0])[1].split("\\")[0]
                    z.write(path, name)
        elif "sharedString" in x:
            for line in filein:
                struct = line.rsplit("sst")
            filein = []
            filein.append(struct[1])
            makeRedactedFile(filein)
            with open(x, "w") as fileout:
                for key in sorted(lines):
                    fileout.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'+struct[0]+'sst '+lines[key]+'sst'+struct[2])
            y = root+"/OUT__"+x.split(root)[1].split("/")[1]+".xlsx"
            z = ZipFile(y, "w")
            zippath = x.rsplit("/xl")
            for ziproot, _, zipfiles in os.walk(zippath[0]):
                for f in zipfiles:
                    path = ziproot+"/"+f
                    if len(path.split(zippath[0])[1].split("\\")) == 3:
                        name = "/"+path.split(zippath[0])[1].split("\\")[1]+"/"+path.split(zippath[0])[1].split("\\")[2]
                    elif len(path.split(zippath[0])[1].split("\\")) == 2:
                        name = "/"+path.split(zippath[0])[1].split("\\")[1]
                    else:
                        name = path.split(zippath[0])[1].split("\\")[0]
                    z.write(path, name)
        else:
            pass
def main():
    global d, password
    def doMatch(md5, proc, path):
        if path.endswith(".csv") or path.endswith(".txt") or path.endswith(".docx") or path.endswith(".xlsx"):
            with open(hashes, "r") as loh:
                for hline in loh:
                    hline = hline.strip("\n")
                    if md5 == hline:
                        stage = "repeat"
                        print(path)
                        doWindow(stage, path)
                        a = r[0]
                        proc[path+"= "+md5] = a
                    else:
                        hashlist.append(md5)
        else:
            pass
    def doHash(d):
        for root, _, files in os.walk(d):
            if len(files) > 0:
                for f in files:
                    print(f)
                    print(root)
                    path = os.path.join(root, f)
                    if not f.startswith("OUT__"):
                        if path.endswith(".csv") or path.endswith(".txt") or path.endswith(".docx") or path.endswith(".xlsx") or path.endswith(".pdf") or path.endswith(".jpg") or path.endswith(".jpeg") or path.endswith(".png") or path.endswith(".gif"):
                            with open(path, "rb") as pf:
                                md5 = hashlib.md5(pf.read()).hexdigest()
                                proc[path+"= "+md5] = "y"
                            doMatch(md5, proc, path)
                        elif path.endswith(".zip"):
                            pass
                        else:
                            stage = "invalid"
                            doWindow(stage, path)
                    else:
                        pass
            else:
                pass
        m_hashes = []
        nm_hashes = []
        for each in hashlist:
            with open(hashes, "r") as hashchk:
                for line in hashchk:
                    if each in line:
                        m_hashes.append(each)
                    else:
                        nm_hashes.append(each)
        mhashset = set(m_hashes)
        nmhashset = set(nm_hashes)
        newhashes = list((set(mhashset)^set(nmhashset)))
        for each in newhashes:
            with open(hashes, "a") as newhash:
                newhash.write(each+"\n")
        for key, value in proc.items():
            if value == "y":
                xd.append(key.split("=")[0])
    d, stage = os.path.realpath(d[0]), "welcome"
    doWindow(stage)
    for root, _, files in os.walk(d):
        if len(files) > 0:
            for f in files:
                if f.endswith(".zip") and not f.startswith("OUT__"):
                    path = os.path.join(root, f)
                    with open(path, "rb") as pf:
                        md5 = hashlib.md5(pf.read()).hexdigest()
                        proc[path+"= "+md5] = "y"
                    doMatch(md5, proc, path)
                    z = ZipFile(path)
                    for info in z.infolist():
                        is_enc = info.flag_bits & 0x1
                    attempt = 0
                    if is_enc:
                        stage = "enc"
                    else:
                        z.extractall(root)
                    while is_enc:
                        doWindow(stage, path, attempt)
                        pswrd = bytes(password[-1].encode("utf-8"))
                        try:
                            z.extractall(root, pwd=pswrd)
                            break
                        except:
                            attempt += 1
                    z.close()
                else:
                    pass
        else:
            pass
    doHash(d)
    for x in xd:
        o = x.split("/")[1]
        if o.startswith("~"):
            stage = "invalid"
            doWindow(stage, x)
        else:
            pass
    for x in xd:
        if x.endswith("bugs") or x.endswith("UserGuide.docx") or x.endswith("redactor.exe") or x.endswith("redactor.py") or x.endswith("hashes") or x.endswith("dictionary") or x.endswith("log"):
            pass
        elif x.endswith(".pdf") or x.endswith(".jpg") or x.endswith(".jpeg") or x.endswith(".png") or x.endswith(".gif"):
            stage = "invalid"
            doWindow(stage, x)
        elif x.endswith(".csv"):
            filetype = "csv"
            with open(x, "r", encoding="utf-8") as filein:
                doPrctvMrkg(root, x, filein, filetype)
        elif x.endswith(".txt"):
            filetype = "txt"
            with open(x, "r", encoding="utf-8") as filein:
                doPrctvMrkg(root, x, filein, filetype)
        elif x.endswith(".docx") or x.endswith(".xlsx"):
            if not x.endswith("UserGuide.docx"):
                filetype, filezip, ziptmp, filein = "office", ZipFile(x), x[:-5], []
                ZipFile.extractall(filezip, ziptmp)
                for ziproot, _, zipfiles in os.walk(ziptmp):
                    for f in zipfiles:
                        path = ziproot+"/"+f
                        if "header" in path or "footer" in path:
                            with open(path, "r", encoding="utf-8") as hf:
                                for line in hf:
                                    if re.findall(re_pm, line):
                                        fields = re.findall(re_pm, line)
                                        for field in fields:
                                            filein.append(field)
                                    else:
                                        pass
                        elif "core" in path:
                            with open(path, "r", encoding="utf-8") as hf:
                                for line in hf:
                                    if re.findall(re_pm, line):
                                        fields = re.findall(re_pm, line)
                                        for field in fields:
                                            filein.append(field)
                                    else:
                                        pass
                        else:
                            pass
                shutil.rmtree(ziptmp)
                filein = set(filein)
                filein = list(filein)
                doPrctvMrkg(root, x, filein, filetype)
            else:
                pass
        else:
            pass
    for x in xd:
        root = x.split("/")[0]
        lines.clear()
        if x.endswith("bugs") or x.endswith("UserGuide.docx") or x.endswith("redactor.exe") or x.endswith("redactor.py") or x.endswith("hashes") or x.endswith("dictionary") or x.endswith("log"):
            pass
        if x.endswith(".csv"):
            filetype = "csv"
            with open(x, "r", encoding="utf-8") as filein:
                doWrite(root, x, filein, filetype)
        elif x.endswith(".txt"):
            filetype = "txt"
            with open(x, "r", encoding="utf-8") as filein:
                doWrite(root, x, filein, filetype)
        elif x.endswith(".docx") or x.endswith(".xlsx"):
            if not x.endswith("UserGuide.docx"):
                filetype = "office"
                filezip = ZipFile(x)
                ziptmp = x[:-5]
                ZipFile.extractall(filezip, ziptmp)
                if x.endswith(".docx"):
                    x = ziptmp+"/word/document.xml"
                    with open(x, "r", encoding="utf-8") as filein:
                        doWrite(root, x, filein, filetype)
                    shutil.rmtree(ziptmp)
                elif x.endswith(".xlsx"):
                    x = ziptmp+"/xl/sharedStrings.xml"
                    with open(x, "r", encoding="utf-8") as filein:
                        doWrite(root, x, filein, filetype)
                    shutil.rmtree(ziptmp)
                else:
                    pass
            else:
                pass
        else:
            pass
    if len(common_words) > 0:
        common_wordsset = set(common_words)
        stage = "review"
        doWindow(stage, common_wordsset)
    else:
        pass
    for x in xd:
        try:
            shutil.os.remove(x)
        except:
            pass
    stage = "finished"
    doWindow(stage)
    sys.exit()
if __name__ == '__main__':
    main()
