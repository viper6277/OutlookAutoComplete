#-*- encoding: utf-8 -*-

"""
About : As noted by : https://support.microsoft.com/en-us/help/2199226

        And with the use of the following software : https://archive.codeplex.com/?p=mfcmapi

        You can extract the Auto Complete data from Outlook 2013, and Outlook 2016...the issue is
        that once you extract the file following the instructions provided.... you get a ton of garbage !!!
        The only useful thing you can do with this is re-import later...but extracting the data....you're left
        SOL !!!

        Not anymore...

"""

__author__ = "Eduard Florea"
__email__ = "viper6277@gmail.com"
__company__ = ""
__copyright__ = "Copyright (C) 2017 {a}".format(a=__author__)
__credits__ = ""
__license__ = "GPLv3"
__version__ = 1.00
__lastdate__ = "2018-02-12"

file_change_log = ''' '''

# ----------------------------------------------------------------------------------------------------------------------
#
# Supporting libraries
#

import csv
import os
import re                         # Need for matching and searching !!!
import sys
import time
import pprint

# ----------------------------------------------------------------------------------------------------------------------
#
# Main Class
#

class OutlookAutoComplete:
    # Class Variables, shared across all instances of the class
    code_version = __version__
    code_author = __author__
    code_company = ''
    code_update_date = __lastdate__
    code_notes = '''This class code is supposed to do lots of cool stuff'''

    code_change_log = '\nChange Log : \n'

    code_change_log += '''
    Updated it on monday !
    '''

    @staticmethod
    def show_credits():
        credits = '\nClass Code Credits : \n'
        credits += '\nAuthor'.ljust(15) + ': %s \n' % OutlookAutoComplete.code_author
        credits += 'Company'.ljust(15) + ': %s \n' % OutlookAutoComplete.code_company
        credits += 'Update On'.ljust(15) + ': %s \n' % OutlookAutoComplete.code_update_date
        credits += 'Notes'.ljust(15) + ': \n'
        credits += '\n' + OutlookAutoComplete.code_notes + '\n'

        return credits

    @staticmethod
    def get_version():
        return 'Class Code Versions : %s' % str(OutlookAutoComplete.code_version)

    @staticmethod
    def change_log():
        return OutlookAutoComplete.code_change_log

    # ------------------------------------------------------------------------------------------------------------------
    #
    # By passing a dictionary to the class constructor....we can plan for more key value pairs in the future.
    #
    def __init__(self, object_config):

        self.file_name = object_config['File Name']

        # This variable will store the contents of the Auto Complete MSG file as a huge string of converted
        # integers... every character get's converted to an ascii decimal integer, so we can do pattern matching on
        # the decimal patterns.
        self.int_string = ''


    def build_int_string(self):
        """
        This method reads the file and converts the byte data to string representations of the byte value!

        """
        with open(self.file_name, "rb") as binary_file:
            # Read the whole file at once
            data = binary_file.read()

        # Uncomment for testing only !
        #print(data)

        # Used for testing only !
        #regular_string = ''

        # Used for testing only !
        #mixed_string = ''

        temp_int_string = ''
        for x in data:

            # Used for testing only !
            #regular_string += str(chr(x))

            # mixed_string += str(chr(x)) + ',' + str(x) + ', '

            # We build a huge string of integers of the data..
            temp_int_string += str(x) + ', '

        self.int_string = temp_int_string


    def extract_data(self, int_string_section):
        """
        This method takes individual integer strings.... and does the data extraction... but it get's called by
        the parse_data() method...

        In researching the raw file....it was evident that there was a marker, that we can use as a seperator...

        before a section of text that includes the name and email... there was a :

           "  n T"

        present.... the spaces are not ASCII 32 as would be expected...but a 0.... and in this space was also a few
        characters that could not be rendered....however when converted to integer a string emerged....:

            110, 0, 221, 1, 15, 84

        110 represents...the letter "n"....and 84 the letter "T" ...the rest was useless....but who else would you
        pattern match....0, 1 and/or 15...??? decimals provided a solution !

        Now we do some fancy find and replace....where we add our own markers...that our parse code will use later.

        We also added spaces after the domain ending....

        """

        '''
        Replace : n T

        with : Name Tag:  
        '''
        int_string_section = int_string_section.replace('110, 0, 221, 1, 15, 84', '78, 97, 109, 101, 32, 84, 97, 103, 58')

        '''
        Replace : SMTP

           with :  Email Tag:  
        '''
        int_string_section = int_string_section.replace('83, 0, 77, 0, 84, 0, 80', '32, 69, 109, 97, 105, 108, 32, 84, 97, 103, 58')

        # --------------------------------------------------------------------------------------------------------------
        #
        # This section can be expanded upon for more domains....
        #

        # add a space at end of a .com
        int_string_section = int_string_section.replace('46, 0, 99, 0, 111, 0, 109', '46, 0, 99, 0, 111, 0, 109, 32')

        # add a space at end of a .net
        int_string_section = int_string_section.replace('46, 0, 110, 0, 101, 0, 116', '46, 0, 110, 0, 101, 0, 116, 32')

        # add a space at end of a .edu
        int_string_section = int_string_section.replace('46, 0, 101, 0, 100, 0, 117', '46, 0, 101, 0, 100, 0, 117, 32')

        # add a space at end of a .gov
        int_string_section = int_string_section.replace('46, 0, 103, 0, 111, 0, 118', '46, 0, 103, 0, 111, 0, 118, 32')

        # --------------------------------------------------------------------------------------------------------------

        # Correct the string, only if it ends with a "," comma.....we eliminate the end comma....
        if int_string_section[-1] == ',':
            int_string_section = int_string_section[:-2]

        temp_list = int_string_section.split(", ")

        num_list = []
        for x in temp_list:
            try:
                x = int(x)
            except ValueError as error:
                pass
            else:
                num_list.append(x)

        # Split the int string into a list, by commas, and convert each to an integer...
        # num_list = [int(x) for x in int_string.split(", ")]

        # --------------------------------------------------------------------------------------------------------------

        converted_string = ''
        for x in num_list:

            # Save only readable characters...
            if x >= 32 and x <= 126:
                converted_string += chr(x)

        converted_string = converted_string.replace('.COM', '.com')
        converted_string = converted_string.replace('.NET', '.net')
        converted_string = converted_string.replace('.GOV', '.gov')
        converted_string = converted_string.replace('.ORG', '.org')
        converted_string = converted_string.replace('.EDU', '.edu')

        # Uncomment for testing only
        # print(converted_string + '\n')

        # --------------------------------------------------------------------------------------------------------------

        name_tag = 'Name Tag:'
        email_tag = 'Email Tag:'

        # Extract Contact Name
        contact_name = converted_string[converted_string.find(name_tag) + len(name_tag): converted_string.find(email_tag)]
        contact_name = contact_name.strip()

        # This can be altered if there is a better way !!!
        email_pattern = r"\b[A-Z0-9._%+-]+(?:@|[(\[]at[\])])[A-Z0-9.-]+\.[A-Z]{2,6}\b"

        r = re.compile(email_pattern, re.IGNORECASE)
        emailAddresses = r.findall(converted_string)

        # --------------------------------------------------------------------------------------------------------------

        if emailAddresses:
            email_address = emailAddresses[0]
        else:
            email_address = ''

        recovered_data = {'Name': contact_name, 'Email Address': email_address}

        return recovered_data


    def parse_data(self):

        self.build_int_string()

        # equal to "n T"...appears before every contact....this can be subject to change
        divider = '110, 0, 221, 1, 15, 84'

        # ----------------------------------------------------------------------------------------------------------------------
        #
        # We use Regular Expressions to search of matches on our divider...
        #


        start_points = []
        for match in re.finditer(divider, self.int_string):

            # Show the matches !
            # print(match.span(), match.group())

            start_points.append(match.span()[0])

        # ----------------------------------------------------------------------------------------------------------------------
        #
        # We now seperate each section into it's own little piece...so we can make it easier to parse...
        #
        data_dict = {}
        previous_start = 0
        i = 0
        for sp in start_points:

            if i >= 1:
                section_data = self.int_string[previous_start: sp]
                data_dict[i] = section_data
                previous_start = sp
            else:
                previous_start = sp

            i += 1

        # Uncomment for testing only !
        # pprint.pprint(data_dict)

        # ----------------------------------------------------------------------------------------------------------------------
        #
        # We now run extraction on every single section...
        #

        raw_list = []
        for k, v in data_dict.items():
            extract_result = self.extract_data(v)

            raw_list.append(extract_result)

        # ----------------------------------------------------------------------------------------------------------------------
        #
        # The raw results we got back, need a little more cleaning...and some duplication removals
        #

        clean_list = {}
        for row in raw_list:
            name = row['Name']
            email = row['Email Address']

            name = name.replace('"', '')
            name = name.replace("'", '')

            email = email.replace('"', '')
            email = email.replace("'", '')

            clean_list[name] = email

        # ----------------------------------------------------------------------------------------------------------------------

        # Uncomment for testing !
        #pprint.pprint(clean_list)

        return clean_list


    def write_to_csv(self, csv_file):

        dict_data = self.parse_data()

        csv_columns = ['Name', 'Email Address']

        try:

            with open(csv_file, 'w', newline='') as outfile:
                writer = csv.DictWriter(outfile, dialect='excel', fieldnames=csv_columns)

                # Write the headers first !!!
                writer.writeheader()

                for k, v in dict_data.items():
                    row_data = {'Name': k, 'Email Address': v}
                    writer.writerow(row_data)

        except IOError as error:
            print(error)

        return


    def __del__(self):
        '''
        Code that executes upon the object deletion
        '''
        pass

# ----------------------------------------------------------------------------------------------------------------------
#
# Test Functions
#

def test_file():

    file_name = 'c:/Maria Auto Complete.msg'

    object_config = {'File Name': file_name}

    x = OutlookAutoComplete(object_config)

    x.write_to_csv('c:/auto.csv')

    #pprint.pprint(x.parse_data())

# ----------------------------------------------------------------------------------------------------------------------
#
# Example Usage
#
'''
'''
# Uncomment for testing only
start_time = time.time()

test_file()

# Uncomment for testing only
print("Function executed in : {t:.15f} seconds".format(t=(time.time() - start_time)))