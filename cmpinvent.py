#!/usr/bin/env python3

# cmpinvent
# Version: 0.1
# Contact: andreaskreisig@gmail.com
# License: MIT

# Copyright (c) 2019, 2020 Andreas Kreisig
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of
# this software and associated documentation files (the "Software"), to deal in
# the Software without restriction, including without limitation the rights to
# use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
# the Software, and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
# FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
# COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
# IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
# CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import argparse
import os
import wmi
import re
import csv
import sys


parser = argparse.ArgumentParser(prog='cmpinvent', description="Creates an inventory of your computers and writes the"
                                                               " collected data into a CSV file.")
parser.add_argument('-p', '--path', type=str,  help="Specify a UNC path and filename to the CSV file.")
parser.add_argument('-m', '--member', type=str, help="Specify Domain or Workgroup name to select a certain NIC on"  # finished
                                                     " a multihomed system.")
parser.add_argument('-d', '--delimiter', type=str, help="Specifies the delimiter used in the CSV file, default is a"  # finished
                                                        " comma.")
parser.add_argument('-c', '--check', help="Check for double entries in the CSV file. Requires either"
                                          " -M or -i option to be set.", action="store_true")
parser.add_argument('-n', '--network', type=str, help="Only consider hosts with a certain network address."
                                                      " Must be given in CIDR notation.")
parser.add_argument('-i', '--index', help="Add an index number to each entry in the CSV file.", action="store_true")  # finished
parser.add_argument('-l', '--location', help="Prompt for the hosts location and add it to the CSV file.",
                    action="store_true")  # finished
parser.add_argument('-I', '--ip', help="Add the IP address to the CSV file.", action="store_true")  # finished
parser.add_argument('-M', '--mac', help="Add the MAC address to the CSV file.", action="store_true")  # finished
parser.add_argument('-H', '--host', help="Add the hostname to the CSV file.", action="store_true")  # finished
parser.add_argument('-C', '--connected', help="Only consider NICs which are connected to a network.",
                    action="store_true")
parser.add_argument('-u', '--user', help="Add current user to CSV file.", action="store_true")
parser.add_argument('-s', '--summary', help="Display a summary after data collection.", action="store_true")
parser.add_argument('-a', '--acknowledge', help="Ask for acknowledge before writing data to CSV file.",
                    action="store_true")
parser.add_argument('-v', '--version', help="Show version information.", action="store_true")  # finished

# parser.add_argument('-l', '--pc-list', metavar='', help="specify a path and file to a host list to remotely request"
#        " the desired information")

c = wmi.WMI()
header = []
system_information = []
physical_nic = []
current_dir = os.getcwd()

args = parser.parse_args()

# Select only physical NICs!
def get_physical_nic():
    for nic in c.Win32_NetworkAdapter(PhysicalAdapter=True):
        physical_nic.append(nic)
    return physical_nic


def get_hostname():
    reg_pattern = re.compile(r'(Name=\")(\S*)(\")')
    wmi_str = str(c.Win32_ComputerSystem(["Name"]))
    hostname = reg_pattern.search(wmi_str).group(2)
    return hostname


def get_macadress():
    nic_list = get_physical_nic()
    for i in nic_list:
        print(i)


def write_csv(header, system_information, delim):
    csv_exists = os.path.isfile(args.path)
    try:
        with open(args.path, mode='a', newline='') as csv_file:
            csv_write = csv.writer(csv_file, delimiter=delim)
            if not csv_exists:
                csv_write.writerow(header)
            csv_write.writerow(system_information)
    except Exception as e:
        print("cmpinvent:", e, file=sys.stderr)


def main():
    delim = ','
    enum = 0
    infodict = {}
    systemdict = {}

    if args.version:
        print("cmpinvent 0.1, copyright (c) by Andreas Kreisig, licensed under the terms of the MIT License.")

    if args.delimiter:
        delim = args.delimiter
        if args.summary:
            infodict['CSV Delimiter'] = delim

    if args.member:
        ident = args.member
        if args.summary:
            infodict['Member'] = ident

    if args.mac:
        get_macadress()

    # This must be the first entry which adds values to the header
    # and system_information list!
    if args.index:
        header.append("#")
        if os.path.isfile(args.path):
            try:
                with open(args.path, "rb") as f:
                    f.seek(-2, os.SEEK_END)
                    while f.read(1) != b"\n":
                        f.seek(-2, os.SEEK_CUR)
                    last_row = f.readline().decode('utf-8').split(delim)
                    var = last_row[0]
                    enum = int(var) + 1
                    system_information.append(enum)
            except Exception as e:
                print("cmpinvent:", e, file=sys.stderr)

        else:
            system_information.append(enum)

    if args.host:
        header.append("Hostname")
        hostname = get_hostname()
        system_information.append(hostname)
        if args.summary:
            systemdict['Hostname'] = hostname

    if args.location:
        header.append("Location")
        site = input("Location: ")
        system_information.append(site)
        if args.summary:
            systemdict['Location: '] = site

    if args.summary:
        if infodict:
            print("\nFilter and CSV setup:")
            for k, v in infodict.items():
                print(k + ":", v)
        if systemdict:
            print("\nSystem information:")
            for k, v in systemdict.items():
                print(k + ":", v)

    # This must always be the last 'if', because it calls the write_csv
    # function hence all information has to be gathered beforehand!
    if args.path:
        write_csv(header, system_information, delim)

    # No flags given
    else:
        print("cmpinvent: for help type 'cmpinvent -h' or 'cmpinvent --help'")
        sys.exit(0)


if __name__ == '__main__':
    main()
