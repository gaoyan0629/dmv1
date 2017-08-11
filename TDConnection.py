import pypyodbc
from util import *

"""
Class Name - CapWords. When using abbreviations in CapWords, \
        capitalize all the letters of the abbreviation. \
        Thus HTTPServerError is better than HttpServerError

Function Name - all lowcase with underscore

Package and Module Name: Modules should have short,\
        all-lowercase names

Method Names and Instance Variables - Use the function naming \
        rules: lowercase with words separated by underscores \
        as necessary to improve readability.

Global Name: The conventions are about the same as those \
        for functions

Exception Name - Because exceptions should be classes, \
        the class naming convention applies here. However,\
        you should use the suffix "Error" on your exception \
        names (if the exception actually is an error).

single_trailing_underscore_ : used by convention to avoid \
        conflicts with Python keyword

Constants are usually defined on a module level and written \
        in all capital letters with underscores separating \
        words. Examples include MAX_OVERFLOW and TOTAL .
"""


class TDConnection():

    def __init__(self, DSN='prod', DBName='bip_vtdb'):
        connectStr = {"DSN": DSN, "DATABASE": DBName}
        self.connection = pypyodbc.connect(**connectStr)

    def fetchall(self, sql_command=""):
        cursor = self.connection.cursor()
        cursor.execute(sql_command)
        retRow = []
        schema_cell = []
        for d in cursor.description:
            if isinstance(d[0], str):
                schema_cell.append(d[0].strip())
            else:
                schema_cell.append(d[0])
        retRow.append(schema_cell)
        for row in cursor.fetchall():
            single_row = []
            for field in row:
                if isinstance(field, str):
                    single_row.append(field.strip())
                else:
                    single_row.append(field)
            retRow.append(single_row)
        return retRow

    def rt(self, sql_command=""):
        pass

    def delete(self, sql_command=""):
        pass

    def close(self, sql_command=""):
        pass

    def __repr__(self):
        pass
# =========================================================================


class ErwinConn():
    def __init__(self, version='v8'):
        self.version = version
        if self.version == 'v8':
            self.connection = pypyodbc.connect(
                    "DSN=ERwin_r8_Current")

    def fetchall(self, sql_command=""):
            cursor = self.connection.cursor()
            cursor.execute(sql_command)
            retRow = []
            schema_cell = []
            for d in cursor.description:
                schema_cell.append(d[0])
            retRow.append(schema_cell)
            for row in cursor.fetchall():
                single_row = []
                for field in row:
                    single_row.append(field)
                retRow.append(single_row)
            return retRow

        def rt(self,sql_command = ""):
            pass

        def delete(self,sql_command = ""):
            pass

        def close(self,sql_command = ""):
            pass

        def __repr__(self):
            pass
