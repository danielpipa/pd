#!/Users/danielpipa/anaconda/bin/python
# -*- coding: utf-8 -*-
"""Process files with Python 3."""

import sys
import subprocess
import time
import datetime
import os
from os.path import join, splitext, isfile, exists
import re
from chardet import utf8prober, latin1prober
import PyPDF2
import filecmp
from send2trash import send2trash
import yaml
import yamlordereddictloader
import yagmail
import unicodedata
import builtins
import shutil
import socket
import glob

if sys.version < '3':
    raise ValueError('Version > 3 required')


# %% Tool classes and functions

def open(*args, **kwargs):
    """Create the directory, if not exists, before opening the file."""
    d, *_ = os.path.split(args[0])
    if exists(d):
        if isfile(d):
            raise TypeError("Directory already exists as a file")
    else:
        os.mkdir(d)

    return builtins.open(*args, **kwargs)


class sp:
    """Useful string processing (sp) functions."""

    @staticmethod  # Doesn't require instantiation
    def quote(s):
        """Quote a string."""
        return '"{}"'.format(s)

    @staticmethod
    def rep_all(s, d):
        """Given a dictionary, perform single character substitutions."""
        return ''.join(d[c] if c in d else c for c in s)

    @staticmethod
    def toisomonth(month):
        """Convert textual to numeral month."""
        mes = {'JAN': '01', 'FEV': '02', 'FEB': '02', 'MAR': '03',
               'ABR': '04', 'APR': '04', 'MAI': '05', 'MAY': '05',
               'JUN': '06', 'JUL': '07', 'AGO': '08', 'AUG': '08',
               'SET': '09', 'SEP': '09', 'OUT': '10', 'OCT': '10',
               'NOV': '11', 'DEZ': '12', 'DEC': '12'}

        try:
            return mes[month.upper()[:3]]
        except KeyError:
            return month

    @staticmethod
    def toisoyear(year):
        ly = len(year)
        if ly == 2:
            return '20' + year
        elif ly == 4:
            return year
        else:
            raise ValueError('Invalid year format')


class pp:
    '''
    Useful path processing (pp) functions.
    '''
    @staticmethod
    def equal(f1, f2):
        if exists(f1) and exists(f2):
            return filecmp.cmp(f1, f2)
        else:
            return False

    @staticmethod
    def check_ren_files(path, add='-', sequential=True):
        '''
        Check if file exists and rename it if necessary.
        '''
        # Avoid passing '', it'd infiniteloop
        i = 1 if sequential else add if add else '-'

        def ren(n):
            name, ext = splitext(n)
            extra = " " if sequential else ""
            return name + extra + str(i) + ext

        path = unicodedata.normalize('NFD', path)  # as Dropbox likes it
        new_path = path
        while exists(new_path) and not pp.equal(path, ren(path)):
            new_path = ren(path)
            i += 1 if sequential else add
        return new_path

    @staticmethod
    def init_tmp_folder(name):
        folder = join(script_folder, name)
        if exists(folder):
            if isfile(folder):
                raise TypeError("Directory already exists as a file")
        else:
            os.mkdir(folder)
        return folder


# %% General classes (any file extension)

class gen:
    '''
    General file class
    The file is moved to a folder named after its extension
    '''

    def init_file(self, path):
        self.path = path  # Full path to file
        self.filefolder, self.filename = os.path.split(path)  # Folder and name
        self.title = splitext(self.filename)[0]  # Name w/o extension
        self.ext = splitext(self.filename)[1].lower()  # Extension

    def __init__(self, path):
        self.init_file(path)
        print(time.strftime('\n%Y-%m-%d %H:%M:%S ') + self.path)

    def identify(self):
        pass

    def move(self, new_filefolder=''):
        if not new_filefolder:
            try:
                new_filefolder = join(self.filefolder, self.move_to_folder)
            except AttributeError:
                new_filefolder = join(self.filefolder, type(self).__name__)

        if not exists(new_filefolder):
            os.mkdir(new_filefolder)

        try:
            new_filename = self.new_filename
        except AttributeError:
            new_filename = self.filename

        new_path = pp.check_ren_files(join(new_filefolder, new_filename))
        os.rename(self.path, new_path)
        self.init_file(new_path)

    def view(self):
        cmd = ['open', sp.quote(self.path)]
        print(subprocess.call(' '.join(cmd), shell=True))


class proc(gen):
    '''
    General class for processing files containing text.
    The file is processed. If identified, it's renamed and moved to process
    folder.
    '''
    def __init__(self, path):
        super().__init__(path)
        self.txt_folder = pp.init_tmp_folder('txt')  # Temporary txt folder
        self.read_yaml_file()

    def move(self):
        if self.proc_sucess:
            for dst in self.regs[self.key][2]:
                if '@' in dst:
                    self.send_email(dst)
                else:
                    super().move(dst)
        else:
            super().move()

    def save_utf8(self, rr):
        '''Save text file in UTF8 format.'''
        old_path = self.path
        self.init_file(join(script_folder, self.new_filename))
        with open(self.path, 'w', encoding='utf-8') as f:
            f.write(self.text)
        send2trash(old_path)
        return True

    def remove_header_empty_lines(self, rr):
        lines = self.text.splitlines()
        i = next(i for i, l in enumerate(lines) if l)
        self.text = '\n'.join(lines[i:])
        return True

    def remove_CR_char(self, rr):
        self.text = self.text.replace('\r', '')
        return True

    def new_filename_gen(self, rr):
        fields = {'date': ('year', 'month', 'day'),
                  'info': ('info{}'.format(i) for i in range(1, 6)),
                  'append': ('append{}'.format(i) for i in range(1, 6))}

        try:
            rr['month'] = sp.toisomonth(rr['month'])  # Convert month to number
        except KeyError:
            pass

        try:
            rr['year'] = sp.toisoyear(rr['year'])
        except KeyError:
            pass

        def get_field_values(key):
            values = [rr[f] for f in fields[key] if f in rr]
            if values == [None]:
                values = ['']
            return values

        self.date = '-'.join(get_field_values('date'))
        info = ' '.join(get_field_values('info'))
        append = ' '.join(get_field_values('append'))

        info = info + ' ' + self.key if info else self.key
        t = '_'.join(i for i in (self.date, info, append) if i)

        self.new_filename = t + self.ext
        return True

    def send_email(self, to):
        try:
            filename = self.new_filename
        except AttributeError:
            filename = self.filename
        to_path = join(to_be_emailed_folder, to + email_addr_token + filename)
        to_path = pp.check_ren_files(to_path)
        shutil.copyfile(self.path, to_path)
        return True

    def split_cards(self, regex_results):
        names = ('MARIA A PIPA', 'KARIM C PIPA')
        ps = '(\s{9}[0-9]\s+?-\s+?%s.+?SubTotal(?:\s+[0-9.]{1,8},[0-9]{2}){2,4})'
        for n in names:
            r = re.findall(ps % n, self.text, re.DOTALL)
            if r:
                filename = '_'.join([self.date, self.key]) + ' ' + n + self.ext
                path = join(self.filefolder, filename)
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(r[0])
                    print("Writing", path)
        return True

    def check_avg_spent(self, regex_results):
        # Calcula média de gastos dos 3 últimos meses dos cartões
        isent = [7300, 8300, 9200, 10200]  # Faixas isenção
        ponts = [9800, 19600, 29400, 39200]  # Pontos a utilizar
        percent = [25, 50, 75, 100]  # Percentuais

        pt = "(?si)Total\s+?da\s+?Fatura\s+?:\s+?R\$\s+?([\d.,]+)"

        files = sorted(glob.glob(self.regs['Visa'][2][0]+"/*.txt"))[-3:]
        files += sorted(glob.glob(self.regs['Master'][2][0]+"/*.txt"))[-3:]
        v = []
        for file in files:
            with open(file, 'r', encoding='utf-8') as f:
                txt = f.read()
                t = re.findall(pt, txt)[0].replace(".", "").replace(",", ".")
                v += [float(t)]

        avgn = sum(v)/3
        atual = sum(1 for i in isent if avgn > i)
        avg = f"{avgn:.2f}"

        name = "danielpipa@gmail.com" + email_addr_token
        name += "Média Cartões " + str(datetime.datetime.now()) + ".txt"
        path = join(to_be_emailed_folder, name)
        with open(path, "w", encoding='utf-8') as f:
            f.write(str(files))
            f.write("\n")
            f.write(avg)
            f.write("\n\n")
            f.write("Tabela isenção\n")
            for i, j in python_zip(isent, percent):
                f.write(f">= {i} -> {j}%\n")
            f.write("\n\n")
            f.write("Situação atual, pontos e isenção\n")
            f.write(f"0 pontos -> {percent[atual-1]}%\n")
            print(atual)
            for i in range(len(isent)-atual):
                print(f"{ponts[i]}\n")
                print(f"{percent[atual+i]}%\n")
                f.write(f"{ponts[i]} pontos -> {percent[atual+i]}%\n")

        return True

    def read_yaml_file(self):
        yaml_file_path = join(script_folder, 'pd.yaml')  # File with regs
        with open(yaml_file_path, encoding='utf8') as f:
            self.regs = yaml.load(f, Loader=yamlordereddictloader.Loader)
        return True

    def identify(self):
        path = join(self.txt_folder, 'txt.txt')
        with open(pp.check_ren_files(path), 'w', encoding='utf-8') as f:
            f.write(self.text)

        for key, value in self.regs.items():
            for reg in value[0]:
                # print(reg)
                # print(self.text)
                r = re.search(reg, self.text)
                if r:
                    r = r.groupdict()
                    self.key = key
                    suc = True
                    for f in value[1]:
                        print(r)
                        # All func's must return true
                        suc &= self.__getattribute__(f)(r)
                    if suc:
                        self.proc_sucess = True
                        return

        print("Couldn't find a regex match")
        self.proc_sucess = False


# %% Classes for specific file types/extensions

class pdf(proc):
    def pdf2txt(self, first_page=1, last_page=10):
        # brew install Caskroom/cask/pdftotext
        prg = '"' + join(script_folder, 'pdftotext') + '"'
        # prg = '/usr/local/bin/pdftotext'
        opts = '-table -f {} -l {} -'.format(first_page, last_page)
        cmd = ' '.join([prg, sp.quote(self.path), opts])
        print(cmd)
        try:
            tmp = subprocess.check_output(cmd, shell=True)
            self.text = tmp.decode('latin1')
        except subprocess.CalledProcessError:
            self.text = ''

    def decrypt(self):
        with open(self.path, 'rb') as f:
            forcepass = False
            try:
                pdf = PyPDF2.PdfFileReader(f)
            except PyPDF2.utils.PdfReadError:
                forcepass = True
            except OSError:
                return
            else:
                if not pdf.isEncrypted:
                    return

        passwords = ['', '998257', '123456', '027']
        # prg = join(script_folder, 'qpdf')  # qpdf installed with homebrew
        prg = '/usr/local/bin/qpdf'  # qpdf installed with homebrew
        opts = '--decrypt --password={pw}'
        dst = join(self.pdf_folder, 'tmp.pdf')
        cmd = ' '.join([prg, opts, sp.quote(self.path), sp.quote(dst)])

        for p in passwords:
            tmp = cmd.format(pw=p)
            print(tmp)
            try:
                subprocess.check_output(tmp, shell=True)
                send2trash(self.path)
                self.init_file(dst)
                break
            except subprocess.CalledProcessError as e:
                if forcepass:
                    send2trash(self.path)
                    self.init_file(dst)
                    break
                if e.returncode != 2:
                    print(e.returncode)
                    raise subprocess.CalledProcessError
                pass

    def __init__(self, path):
        super().__init__(path)
        self.pdf_folder = pp.init_tmp_folder('pdf')  # Temporary pdf folder
        self.decrypt()
        self.pdf2txt()


class txt(proc):
    def __init__(self, path):
        super().__init__(path)
        with open(self.path, 'rb') as f:
            tmp = f.read()

        u8 = utf8prober.UTF8Prober()
        u8.feed(tmp)
        l1 = latin1prober.Latin1Prober()
        l1.feed(tmp)

        # Detect encoding
        if u8.get_confidence() > l1.get_confidence():
            enc = 'utf-8'
        else:
            enc = 'latin1'

        with open(self.path, encoding=enc, newline='\r\n') as f:
            self.text = sp.rep_all(f.read(), {chr(160): ' ',
                                   '\r': '', chr(0): ''})


class epub(proc):
    def epub2txt(self, opts=''):
        prg = '/Applications/calibre.app/Contents/MacOS/ebook-convert'
        dst = pp.check_ren_files(join(self.txt_folder, 'txt.txt'))
        cmd = ' '.join([prg, sp.quote(self.path), sp.quote(dst), opts])
        subprocess.check_call(cmd, shell=True)
        with open(dst, 'r', encoding='utf8') as fd:
            self.text = fd.read()

    def __init__(self, path):
        super().__init__(path)
        self.epub2txt()

# %% File extensions that do not need processing, just move to folder


class docx(gen):
    pass


class doc(gen):
    move_to_folder = 'docx'


class dot(gen):
    move_to_folder = 'docx'


class docm(gen):
    move_to_folder = 'docx'


class rtf(gen):
    move_to_folder = 'docx'


class xlsx(gen):
    pass


class xls(gen):
    move_to_folder = 'xlsx'


python_zip = zip


class zip(gen):
    move_to_folder = 'archive'


class rar(gen):
    move_to_folder = 'archive'


class m(gen):
    pass


class xml(gen):
    pass


class dmg(gen):
    pass


class mbz(gen):
    move_to_folder = 'moodle'


class pptx(gen):
    pass


class ppsx(gen):
    move_to_folder = 'pptx'


class ppt(gen):
    move_to_folder = 'pptx'


class pps(gen):
    move_to_folder = 'pptx'


class png(gen):
    move_to_folder = 'figures'


class jpg(gen):
    move_to_folder = 'figures'


class jpeg(gen):
    move_to_folder = 'figures'


class mp4(gen):
    move_to_folder = 'videos'


class torrent(gen):
    pass


# %%


class logger():
    def __init__(self, std, path):
        self.terminal = std
        self.log = open(path, 'a', encoding='utf-8')

    def write(self, msg):
        self.terminal.write(msg)
        self.log.write(msg)

    def flush(self):
        pass


def email_files():
    # TODO: Para avisar se chega o email mensal do condomínio, Net, e outras
    # coisas que se paga por email.
    # TODO: Verificar se existe o arquivo no Dropbox do boleto referente

    # Try to email files
    for file in os.listdir(to_be_emailed_folder):
        path1 = join(to_be_emailed_folder, file)
        r = re.match('(.+?@.+)' + email_addr_token + '(.+)', file)
        if r:
            to = r.group(1)
            filename = r.group(2)
            path2 = join(to_be_emailed_folder, filename)
            shutil.copyfile(path1, path2)
            try:
                yag = yagmail.SMTP('danielpipa', 'vzwekdvtetxavaoi')
                yag.useralias = 'Daniel Rodrigues Pipa <danielpipa@gmail.com>'
                yag.send(to, filename, path2, bcc='drpipa@hotmail.com')
                send2trash(path1)
            except socket.gaierror:
                pass
            finally:
                os.remove(path2)
        else:
            send2trash(path1)


if __name__ == '__main__':

    script_folder = os.path.dirname(os.path.realpath(__file__))
    to_be_emailed_folder = join(script_folder, 'to_be_emailed')
    if not os.path.isdir(to_be_emailed_folder):
        os.mkdir(to_be_emailed_folder)
    email_addr_token = '_@_@@_'  # Token to codify email destiny in filename

    sys.stderr = logger(sys.stderr, join(script_folder, 'stderr.txt'))
    sys.stdout = logger(sys.stdout, join(script_folder, 'stdout.txt'))
    watched_folder = '/Users/danielpipa/Downloads'

    for file in next(os.walk(watched_folder))[2]:
        ext = splitext(file)[1].lower()
        path = join(watched_folder, file)
        try:
            cls = globals()[ext[1:]]
            obj = cls(path)
        except (KeyError, FileNotFoundError):
            continue

        obj.identify()
        obj.move()
        obj.view()

    email_files()
