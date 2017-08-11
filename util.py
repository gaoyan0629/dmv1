from __future__ import print_function
import os
import sys
import re
import csv
import datetime
import subprocess
import shutil

# ============================================================


def validate_schema(source, target):
    """ this checks if the value defined in source are \
            in target or not. if found, it also put the\
            position (pos)
            """
    ret = []
    for i in range(len(source)):
        found = False
        pos = None
        for k in range(len(target)):
            if (
                    (
                        isinstance(source[i], unicode)
                        or isinstance(source[i], str)
                    )
                    and (
                        isinstance(target[k], unicode)
                        or isinstance(target[k], str)
                        )
               ):
                if source[i].strip().lower() == \
                        target[k].strip().lower():
                    found = True
                    pos = k
                    break
        ret.append([source[i], pos, found])
    return ret
# =========================================================


def walk(dirname, pattern):
	"""
        this is the customized version of python walk,\
                it search the root directory of dirname\
                based on the pattern
        """
        retfile = []
        regex = re.compile(pattern)
        for path, subdirs, files in os.walk(dirname):
                for f in files:
                        if regex.search(f):
                                file_path = os.path.join(
                                        os.path.abspath(path),
                                        f)
                                retfile.append([file_path,f])
        return retfile
# =============================================================


def dedup(data, keyCol):
        """
        this will de-duplicate from the list in data based\
                on the key defined in data col
        """
        sourcedict = {}
        retlist = []
        for i in range(len(data)):
                sourcetemp = ""
                for j in keyCol:
                        sourcetemp += data[i][j-1]
                        sourcetemp += '||'
                sourcedict.setdefault(sourcetemp, None)

                if not sourcedict[sourcetemp]:
                        sourcedict[sourcetemp] = [i, data[i]]
        for key, value in sorted(
                sourcedict.iteritems(),
                key=lambda (k, v): (v[0], k)):
                retlist.append(value[1])

        return retlist

# ===========================================================


def lookup(source, lookup, sourcecol, lookupcol):
        """
        this will look up from list in (lookup) and return \
                its position in relation to where it lookup
        """
        sourcekey = []
        lookupkey = []

        for i in range(len(source)):
                sourcetemp = ""
                for j in sourcecol:
                        sourcetemp += source[i][j-1]
                        sourcetemp += '||'
                sourcekey.append(sourcetemp)
        for i in range(len(lookup)):
                lookuptemp = ""
                for j in lookupcol:
                        lookuptemp += lookup[i][j-1]
                        lookuptemp += '||'
                lookupkey.append(lookuptemp)

        foundkey = []

        for i in range(len(source)):
                found = False
                pos = None
                for j in range(len(lookup)):
                        if sourcekey[i] == lookupkey[j]:
                                found = True
                                pos = j
                                break
                foundkey.append([found, pos])
        return foundkey

# =============================================================


def csv_writer(data, fileName, delimiter=','):
    """
    Write data to a CSV file
    """
    with open(fileName, "wb") as csv_file:
        writer = csv.writer(csv_file, delimiter=delimiter)
        for line in data:
                line = [str(s).encode('utf-8') for s in line]
                writer.writerow(line)
    return True

# =============================================================


def csv_dict_writer(data, fileName, fieldnames, delimiter=','):
    """
    Writes a CSV file using DictWriter
    """
    with open(fileName, "wb") as out_file:
        writer = csv.DictWriter(
                out_file,
                delimiter=delimiter,
                fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow(row)
    return True
# =============================================================


def csv_reader(fileName, delimiter=','):
    """
    Read a csv file
    """
    data = []
    with open(fileName, 'rb') as csvfile:
        reader = csv.reader(csvfile, delimiter=delimiter)
        for row in reader:
            data.append(row)
    return data
# =============================================================


def remove_unicode(data):
        regex = re.compile(r"""([^\x00-\x7F])""")
        ret = []
        for i in range(len(data)):
                row = []
                for j in range(len(data[i])):
                    if isinstance(
                            data[i][j], unicode) or \
                                    isinstance(data[i][j], str):
                                        row.append(
                                                regex.sub(
                                                    '',
                                                    data[i][j]))
                    else:
                        row.append(data[i][j])
                ret.append(row)
        return ret

# =============================================================


def eprint(*args, **kwargs):
    print('[' + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ']' + '[ERROR]', *args, file=sys.stderr, **kwargs)


def iprint(*args, **kwargs):
    print(
            '[' + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ']' + '[INFO]', *args, **kwargs)

# =============================================================


class ExcelHeaderError(Exception):
    "this is exception when the head of excel file is wrong"
    pass


class fileSuffixError(Exception):
    """this is exception when the file suffix of excel file \
            is wrong
    """
    pass


class NoSupportError(Exception):
    """this is exception when the using the feature that have \
            not been supported
    """
    pass


def getDataTypeAndLen(string):
    charPattern = r"""(\w*?char).*?\((\d+).*\)"""
    numericPattern = r"""(nume.*?).*?\((\d+).*\)"""
    dateTimePattern = r"""(date[\w]?)"""
    decimalPattern = r"""(decimal).*?\((\d+).*\)"""
    bigintPattern = r"""bigint|smallint|integer"""
    charReg = re.compile(charPattern, re.IGNORECASE)
    data = charReg.findall(string)
    if data:
        return data[0][0], data[0][1]
    numericReg = re.compile(numericPattern, re.IGNORECASE)
    data = numericReg.findall(string)
    if data:
        return data[0][0], data[0][1]
    decimalReg = re.compile(decimalPattern, re.IGNORECASE)
    data = decimalReg.findall(string)
    if data:
        return data[0][0], data[0][1]
    dateTimeReg = re.compile(dateTimePattern, re.IGNORECASE)
    data = dateTimeReg.findall(string)
    if data:
        return data[0], ""
    bigintReg = re.compile(bigintPattern, re.IGNORECASE)
    data = bigintReg.findall(string)
    if data:
        return data[0], ""
    return "", ""


def getDataSourceCd(string):
    pattern = r"""(\d{3})"""
    reg = re.compile(pattern, re.IGNORECASE)
    data = reg.findall(string)
    if data:
        return data[0]
    else:
        return ""


class GenSeq():
    count = 0

    @staticmethod
    def genseq():
        GenSeq.count += 1
        return GenSeq.count

    @staticmethod
    def attachseq(List):
        GenSeq.count += 1
        List.insert(0, GenSeq.count)


def netDriveConn():
    if os.path.exists("""Y:\\"""):
        subprocess.call(r"""net use Y: /delete""")
        subprocess.call(
                r"""net use Y: "http://teamsite.1bank.dbs.com/sites/EnterpriseServices/BDWAnalyst/BIP Design Documents/CDM Design Documents" """, shell=True)

    if os.path.exists("""Z:\\"""):
        subprocess.call(r"""net use Z: /delete""")
        subprocess.call(r"""net use Z: "http://teamsite.1bank.dbs.com/sites/EnterpriseServices/BDWAnalyst/BIP Design Documents/OIA Details" """, shell=True)


def copytree(src, dst, symlinks=False, ignore=None):
    if not os.path.exists(dst):
        os.makedirs(dst)
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            copytree(s, d, symlinks, ignore)
        else:
            if not os.path.exists(d) or os.stat(s).st_mtime - os.stat(d).st_mtime > 1:
                shutil.copy2(s, d)
