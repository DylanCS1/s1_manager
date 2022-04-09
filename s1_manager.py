""" s1_manager.py 
    License: MIT license
    Source: https://github.com/DylanCS1/s1_manager

"""

import asyncio
import csv
import datetime
import json
import logging
import os
import time
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext as ScrolledText
from functools import partial
from tkinter import ttk

import aiohttp
import requests
from xlsxwriter.workbook import Workbook

import gui

# Consts #############################
__version__ = "2022.0.2"
api_version = "v2.1"
dir_path = os.path.dirname(os.path.realpath(__file__))

window = tk.Tk()
window.title("S1 Manager")
window.minsize(900, 700)

window.tk.call("source", os.path.join(dir_path, ".THEME/forest-dark.tcl")) # https://github.com/rdbende/Forest-ttk-theme
ttk.Style().theme_use("forest-dark")

loginMenuFrame = tk.Frame()
mainMenuFrame = tk.Frame()
exportFromDVFrame = tk.Frame()
upgradeFromCSVFrame = tk.Frame()
exportActivityLogFrame = tk.Frame()
moveAgentsFrame = tk.Frame()
assignCustomerIdentifierFrame = tk.Frame()
decomissionAgentsFrame = tk.Frame()
exportAllAgentsFrame = tk.Frame()
error = tk.StringVar()
hostname = tk.StringVar()
apitoken = tk.StringVar()
proxy = tk.StringVar()
inputcsv = tk.StringVar()
useSSL = tk.BooleanVar()
useSSL.set(True)
exportExclusionsFrame = tk.Frame()


class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget
    Adapted from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06"""

    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state="normal")
            self.text.insert(tk.END, msg + "\n")
            self.text.configure(state="disabled")
            # Autoscroll to the bottom
            self.text.yview(tk.END)

        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)


def testLogin(hostname, apitoken, proxy):
    headers = {
        "Content-type": "application/json",
        "Authorization": "ApiToken " + apitoken,
    }
    r = requests.get(
        hostname + f"/web/api/{api_version}/system/info",
        headers=headers,
        proxies={"http": proxy, "https": proxy},
        verify=useSSL.get(),
    )
    if r.status_code == 200:
        return headers, True
    elif r.status_code != 200:
        headers = {
            "Content-type": "application/json",
            "Authorization": "Token " + apitoken,
        }
    r = requests.get(
        hostname + f"/web/api/{api_version}/system/info",
        headers=headers,
        proxies={"http": proxy, "https": proxy},
        verify=useSSL.get(),
    )
    if r.status_code == 200:
        return headers, True
    else:
        return 0, False


def login():
    hostname.set(consoleAddressEntry.get())
    apitoken.set(apikTokenEntry.get())
    proxy.set(proxyEntry.get())
    global headers
    headers, login_succ = testLogin(hostname.get(), apitoken.get(), proxy.get())
    if login_succ:
        loginMenuFrame.pack_forget()
        mainMenuFrame.pack()
    else:
        tk.Label(
            master=loginMenuFrame,
            text="Login to the management console failed. Please check your credentials and try again",
            fg="red",
        ).grid(row=9, column=0, columnspan=2, pady=10)


def goBacktoMainPage():
    _list = window.winfo_children()
    for item in _list:
        if item.winfo_children():
            _list.extend(item.winfo_children())
    for item in _list:
        if isinstance(item, tk.Toplevel) is not True:
            item.pack_forget()
    mainMenuFrame.pack()


def switchFrames(framename):
    mainMenuFrame.pack_forget()
    framename.pack()


def exportFromDV():
    async def dv_query_to_csv(
        querytype, session, hostname, dv_query_id, headers, firstrun, proxy
    ):
        params = f"/web/api/{api_version}/dv/events/{querytype}?queryId={dv_query_id}"
        url = hostname + params
        while url:
            async with session.get(
                url, headers=headers, proxy=proxy, ssl=useSSL.get()
            ) as response:
                if response.status != 200:
                    error = (
                        f"Status: {str(response.status)} Problem with the request. Exiting."
                    )
                    tk.Label(master=exportFromDVFrame, text=error, fg="red").grid(
                        row=6, column=0, pady=2
                    )
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for data in data:
                            if querytype == "file":
                                f = csv.writer(
                                    open(
                                        "dv_file.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "ip":
                                f = csv.writer(
                                    open(
                                        "dv_ip.csv", "a+", newline="", encoding="utf-8"
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "url":
                                f = csv.writer(
                                    open(
                                        "dv_url.csv", "a+", newline="", encoding="utf-8"
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "dns":
                                f = csv.writer(
                                    open(
                                        "dv_dns.csv", "a+", newline="", encoding="utf-8"
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "process":
                                f = csv.writer(
                                    open(
                                        "dv_process.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "registry":
                                f = csv.writer(
                                    open(
                                        "dv_registry.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "scheduled_task":
                                f = csv.writer(
                                    open(
                                        "dv_scheduled_task.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)
                    if cursor:
                        paramsnext = (
                            f"/web/api/{api_version}/dv/events/{querytype}?cursor={cursor}&queryId={dv_query_id}&limit=100"
                        )
                        url = hostname + paramsnext
                    else:
                        url = None

    async def run(hostname, dv_query_id, apitoken, proxy):
        async with aiohttp.ClientSession() as session:
            for query in dv_query_id:
                firstrun = False
                if query == dv_query_id[0]:
                    firstrun = True
                typefile = asyncio.create_task(
                    dv_query_to_csv(
                        "file", session, hostname, query, headers, firstrun, proxy
                    )
                )
                typeip = asyncio.create_task(
                    dv_query_to_csv(
                        "ip", session, hostname, query, headers, firstrun, proxy
                    )
                )
                typeurl = asyncio.create_task(
                    dv_query_to_csv(
                        "url", session, hostname, query, headers, firstrun, proxy
                    )
                )
                typedns = asyncio.create_task(
                    dv_query_to_csv(
                        "dns", session, hostname, query, headers, firstrun, proxy
                    )
                )
                typeprocess = asyncio.create_task(
                    dv_query_to_csv(
                        "process", session, hostname, query, headers, firstrun, proxy
                    )
                )
                typeregistry = asyncio.create_task(
                    dv_query_to_csv(
                        "registry", session, hostname, query, headers, firstrun, proxy
                    )
                )
                typescheduledtask = asyncio.create_task(
                    dv_query_to_csv(
                        "scheduled_task",
                        session,
                        hostname,
                        query,
                        headers,
                        firstrun,
                        proxy,
                    )
                )
                await typefile
                await typeip
                await typeurl
                await typedns
                await typeprocess
                await typeregistry
                await typescheduledtask

    dv_query_id = queryIdEntry.get()
    if dv_query_id:
        dv_query_id = dv_query_id.split(",")
        asyncio.run(run(hostname.get(), dv_query_id, apitoken.get(), proxy.get()))
        filename = "-"
        filename = filename.join(dv_query_id)
        workbook = Workbook(filename + ".xlsx")
        csvs = [
            "dv_file.csv",
            "dv_ip.csv",
            "dv_url.csv",
            "dv_dns.csv",
            "dv_process.csv",
            "dv_registry.csv",
            "dv_scheduled_task.csv",
        ]
        for csvfile in csvs:
            worksheet = workbook.add_worksheet(csvfile.split(".")[0])
            if os.path.isfile(csvfile):
                with open(csvfile, "r", encoding="utf8") as f:
                    reader = csv.reader(f)
                    for r, row in enumerate(reader):
                        for c, col in enumerate(row):
                            worksheet.write(r, c, col)
                os.remove(csvfile)
        workbook.close()
        done = f"Done! Created the file {filename}.xlsx"
        tk.Label(master=exportFromDVFrame, text=done).grid(
            row=6, column=0, pady=2
        )
    else:
        tk.Label(
            master=exportFromDVFrame,
            text="No DV Query ID found. Please try again",
            fg="red",
        ).grid(row=6, column=0, pady=2)


def exportActivityLog(searchOnly):
    st = ScrolledText.ScrolledText(master=exportActivityLogFrame, state="disabled", height=10)
    st.configure(font="TkFixedFont")
    st.grid(row=11, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="activitylogexport.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    os.environ["TZ"] = "UTC"
    p = "%Y-%m-%d"
    fromdate_epoch = str(int(time.mktime(time.strptime(dateFrom.get(), p)))) + "000"
    todate_epoch = str(int(time.mktime(time.strptime(dateTo.get(), p)))) + "000"
    if dateFrom.get() and dateTo.get():
        url = (
            hostname.get()
            + f"/web/api/{api_version}/activities?limit=1000&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&includeHidden=false"
        )
        if searchOnly:
            while url:
                response = requests.get(
                    url,
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        f"Status: {str(response.status_code)} Problem with the request. Details - {str(response.text)}"
                    )
                else:
                    data = response.json()
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for item in data:
                            if (
                                stringSearchEntry.get().upper()
                                in item["primaryDescription"].upper()
                            ):
                                logger.info(
                                    f"{item['createdAt']} - {item['primaryDescription']} - {item['secondaryDescription']}"
                                )
                            elif item["secondaryDescription"]:
                                if (
                                    stringSearchEntry.get().upper()
                                    in item["secondaryDescription"].upper()
                                ):
                                    logger.info(
                                        f"{item['createdAt']} - {item['primaryDescription']} - {item['secondaryDescription']}"
                                    )
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/activities?limit=1000&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&cursor={cursor}&includeHidden=false"
                        url = hostname.get() + paramsnext
                    else:
                        url = None
        else:
            timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            f = csv.writer(
                open(
                    f"activityLogExport{timestamp}.csv",
                    "a+",
                    newline="",
                    encoding="utf-8",
                )
            )
            firstrun = True
            while url:
                response = requests.get(
                    url,
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        f"Status: {str(response.status_code)} Problem with the request. Details - {str(response.text)}"
                    )
                else:
                    data = response.json()
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        if firstrun:
                            tmp = []
                            for key, value in data[0].items():
                                tmp.append(key)
                            f.writerow(tmp)
                            firstrun = False
                        for item in data:
                            tmp = []
                            for key, value in item.items():
                                tmp.append(value)
                            f.writerow(tmp)
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/activities?limit=1000&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&cursor={cursor}&includeHidden=false"
                        url = hostname.get() + paramsnext
                    else:
                        url = None
            logger.info(f"Done! Output file is - activityLogExport{timestamp}.csv")

    else:
        logger.error("You must state a FROM date and a TO date")


def upgradeFromCSV(justPackages):
    st = ScrolledText.ScrolledText(master=upgradeFromCSVFrame, state="disabled", height=10)
    st.configure(font="TkFixedFont")
    st.grid(row=11, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="upgradefromcsv.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    if justPackages:
        params = f"/web/api/{api_version}/update/agent/packages?sortBy=updatedAt&sortOrder=desc&countOnly=false&limit=1000"
        url = hostname.get() + params
        f = csv.writer(open("packages_list.csv", "a+", newline="", encoding="utf-8"))
        f.writerow(
            [
                "Name",
                "ID",
                "Version",
                "OS Type",
                "OS Arch",
                "Package Type",
                "File Extension",
                "Status",
                "Scope Level",
            ]
        )

        while url:
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    f"Status: {str(response.status_code)} Problem with the request. Details - {str(response.text)}"
                )
            else:
                data = response.json()
                cursor = data["pagination"]["nextCursor"]
                data = data["data"]
                if data:
                    for data in data:
                        f.writerow(
                            [
                                [data["fileName"]],
                                data["id"],
                                data["version"],
                                data["osArch"],
                                data["osType"],
                                data["packageType"],
                                data["fileExtension"],
                                data["status"],
                                data["scopeLevel"],
                            ]
                        )
                if cursor:
                    paramsnext = (
                        f"/web/api/{api_version}/update/agent/packages?sortBy=updatedAt&sortOrder=desc&limit=1000&cursor={cursor}&countOnly=false"
                    )
                    url = hostname.get() + paramsnext
                else:
                    url = None
        logger.info("Printed packages list into packages_list.csv")
    else:
        with open(inputcsv.get()) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            line_count = 0
            for row in csv_reader:
                logger.info(f"\t Upgrading endpoint named -  {row[0]}")
                url = hostname.get() + f"/web/api/{api_version}/agents/actions/update-software"
                body = {
                    "filter": {"computerName": row[0]},
                    "data": {"packageId": packageIDEntry.get()},
                }
                response = requests.post(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        f"Failed to upgrade endpoint {row[0]} Error code: {str(response.status_code)} Description: {str(response.text)}"
                    )
                else:
                    data = response.json()
                    logger.info(
                        f'Sent upgrade command to {data["data"]["affected"]} endpoints'
                    )
                line_count += 1
            logger.info(f"Finished! Processed {line_count} lines.")


def moveAgents(justGroups):
    st = ScrolledText.ScrolledText(master=moveAgentsFrame, state="disabled", height=10)
    st.configure(font="TkFixedFont")
    st.grid(row=8, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="moveagentsfromcsv.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    if justGroups:
        params = (
            "/web/api/v2.0/groups?isDefault=false&limit=100&type=static&countOnly=false" #TODO 2.0?
        )
        url = hostname.get() + params
        f = csv.writer(open("group_to_id_map.csv", "a+", newline="", encoding="utf-8"))
        f.writerow(["Name", "ID", "Site ID", "Created By"])
        while url:
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    f"Status: {str(response.status_code)} Problem with the request. Details - {str(response.text)}"
                )
            else:
                data = response.json()
                cursor = data["pagination"]["nextCursor"]
                data = data["data"]
                if data:
                    for data in data:
                        f.writerow(
                            [
                                [data["name"]],
                                data["id"],
                                data["siteId"],
                                data["creator"],
                            ]
                        )
                if cursor:
                    paramsnext = (
                        f"/web/api/v2.0/groups?isDefault=false&limit=100&type=static&cursor={cursor}&countOnly=false" #TODO 2.0?
                    )
                    url = hostname.get() + paramsnext
                else:
                    url = None
        logger.info("Added group mapping to the file group_to_id_map.csv ")
    else:
        with open(inputcsv.get()) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            line_count = 0
            for row in csv_reader:
                logger.info(f"\t Moving endpoint name {row[0]} to Site ID {row[2]}")
                url = hostname.get() + f"/web/api/{api_version}/agents/actions/move-to-site"
                body = {
                    "filter": {"computerName": row[0]},
                    "data": {"targetSiteId": row[2]},
                }
                response = requests.post(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        f"Failed to transfer endpoint {row[0]} to site {row[1]} Error code: {str(response.status_code)} Description: {str(response.text)}"
                    )
                else:
                    data = response.json()
                    logger.info(f'Moved {data["data"]["affected"]} endpoints')
                logger.info(f"\t Moving endpoint name {row[0]} to Group ID {row[1]}")
                url = hostname.get() + f"/web/api/{api_version}/groups/" + row[1] + "/move-agents"
                body = {"filter": {"computerName": row[0]}}
                response = requests.put(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        f"Failed to transfer endpoint {row[0]} to group {row[1]} Error code: {str(response.status_code)} Description: {str(response.text)}"
                    )
                else:
                    data = response.json()
                    logger.info(f'Moved {data["data"]["agentsMoved"]} endpoints')
                line_count += 1
            logger.info(f"Finished! Processed {line_count} lines.")


def assignCustomerIdentifier():
    st = ScrolledText.ScrolledText(
        master=assignCustomerIdentifierFrame, state="disabled", height=10
    )
    st.configure(font="TkFixedFont")
    st.grid(row=8, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="upgradefromcsv.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    with open(inputcsv.get()) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0
        for row in csv_reader:
            logger.info(f"\t Updating customer identifier for endpoint -  {row[0]}")
            url = hostname.get() + f"/web/api/{api_version}/agents/actions/set-external-id"
            body = {
                "filter": {"computerName": row[0]},
                "data": {"externalId": customerIdentifierEntry.get()},
            }
            response = requests.post(
                url,
                data=json.dumps(body),
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    f"Failed to update customer identifier for endpoint {row[0]} Error code: {str(response.status_code)} Description: {str(response.text)}"
                )
            else:
                r = response.json()
                affected_num_of_endpoints = r["data"]["affected"]
                if affected_num_of_endpoints < 1:
                    logger.info(f"No endpoint matched the name {row[0]}")
                elif affected_num_of_endpoints > 1:
                    logger.info(
                        f"{affected_num_of_endpoints} endpoints matched the name {row[0]} , customer identifier was updated for all"
                    )
                else:
                    logger.info(f"Successfully updated the customer identifier")
            line_count += 1
        logger.info(f"Finished! Processed {line_count} lines.")


def exportAllAgents():
    st = ScrolledText.ScrolledText(master=exportAllAgentsFrame, state="disabled", height=10)
    st.configure(font="TkFixedFont")
    st.grid(row=4, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="exportallagentstocsv.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    f = csv.writer(
        open(f"endpointsexport_{timestamp}.csv", "a+", newline="", encoding="utf-8")
    )
    firstrun = True
    url = hostname.get() + f"/web/api/{api_version}/agents?limit=100"
    while url:
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": proxy.get(), "https": proxy.get()},
            verify=useSSL.get(),
        )
        if response.status_code != 200:
            logger.error(
                f"Status: {str(response.status_code)} Problem with the request. Details - {str(response.text)}"
            )
        else:
            data = response.json()
            cursor = data["pagination"]["nextCursor"]
            data = data["data"]
            if data:
                if firstrun:
                    tmp = []
                    for key, value in data[0].items():
                        tmp.append(key)
                    f.writerow(tmp)
                    firstrun = False
                for item in data:
                    tmp = []
                    for key, value in item.items():
                        tmp.append(value)
                    f.writerow(tmp)
            if cursor:
                paramsnext = f"/web/api/{api_version}/agents?limit=100&cursor={cursor}"
                url = hostname.get() + paramsnext
            else:
                url = None
    logger.info(f"Done! Output file is - endpointsexport_{timestamp}.csv")


def decomissionAgents():
    st = ScrolledText.ScrolledText(master=decomissionAgentsFrame, state="disabled", height=10)
    st.configure(font="TkFixedFont")
    st.grid(row=6, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="decomissionagentfromcsv.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    with open(inputcsv.get()) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0
        for row in csv_reader:
            logger.info(f"\t Decomissioning Endpoint -  {row[0]}")
            logger.info(f"Getting endpoint ID for {row[0]}")
            url = (
                hostname.get()
                + f"/web/api/{api_version}/agents?countOnly=false&computerName={row[0]}&limit=1000"
            )
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    f"Failed to get ID for endpoint {row[0]} Error code: {str(response.status_code)} Description: {str(response.text)}"
                )
            else:
                r = response.json()
                totalitems = r["pagination"]["totalItems"]
                if totalitems < 1:
                    logger.info(
                        f"Could not locate any IDs for endpoint named {row[0]} - Please note the query is CaSe SenSiTiVe"
                    )
                else:
                    r = r["data"]
                    uuidslist = []
                    for item in r:
                        uuidslist.append(item["id"])
                        logger.info(
                            f"Found ID {item['id']}! adding it for decomissining"
                        )
                    url = hostname.get() + f"/web/api/{api_version}/agents/actions/decommission"
                    body = {"filter": {"ids": uuidslist}}
                    response = requests.post(
                        url,
                        data=json.dumps(body),
                        headers=headers,
                        proxies={"http": proxy.get(), "https": proxy.get()},
                        verify=useSSL.get(),
                    )
                    if response.status_code != 200:
                        logger.error(
                            f"Failed to decomission endpoint {row[0]} Error code: {str(response.status_code)} Description: {str(response.text)}"
                        )
                    else:
                        r = response.json()
                        affected_num_of_endpoints = r["data"]["affected"]
                        if affected_num_of_endpoints < 1:
                            logger.info(f"No endpoint matched the name {row[0]}")
                        elif affected_num_of_endpoints > 1:
                            logger.info(
                                f"{affected_num_of_endpoints} endpoints matched the name {row[0]} , all of them got decomissioned"
                            )
                        else:
                            logger.info(f"Successfully decomissioned the endpoint")
            line_count += 1
        logger.info(f"Finished! Processed {line_count} lines.")


def exportExclusions():
    st = ScrolledText.ScrolledText(master=exportExclusionsFrame, state="disabled", height=10)
    st.configure(font="TkFixedFont")
    st.grid(row=4, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename="exportExclusions.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    async def getAccounts(session):
        params = (
            f"/web/api/{api_version}/accounts?limit=100" + "&countOnly=false&tenant=true"
        )
        url = hostname.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                if response.status != 200:
                    logger.error(
                        f"Status: {str(response.status)} Problem with the request. Exiting."
                    )
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for account in data:
                            # print('ACCOUNT: ' + account['id'] + ' | ' + account['name'])
                            dictAccounts[account["id"]] = account["name"]
                    if cursor:
                        paramsnext = (
                            f"/web/api/{api_version}/accounts?limit=100&cursor={cursor}&countOnly=false&tenant=true"
                        )
                        url = hostname.get() + paramsnext
                    else:
                        url = None

    async def getSites(session):
        params = (
            f"/web/api/{api_version}/sites?limit=100&countOnly=false&tenant=true"
        )
        url = hostname.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                if response.status != 200:
                    logger.error(
                        f"Status: {str(response.status)} Problem with the request. Exiting."
                    )
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for site in data["sites"]:
                            # print('SITE: ' + site['id'] + ' | ' + site['name'])
                            dictSites[site["id"]] = site["name"]
                    if cursor:
                        paramsnext = (
                            f"/web/api/{api_version}/sites?limit=100&cursor={cursor}&countOnly=false&tenant=true"
                        )
                        url = hostname.get() + paramsnext
                    else:
                        url = None

    async def getGroups(session):
        params = (
            f"/web/api/{api_version}/groups?limit=100&countOnly=false&tenant=true"
        )
        url = hostname.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                if response.status != 200:
                    logger.error(
                        f"Status: {str(response.status)} Problem with the request. Exiting."
                    )
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for group in data:
                            # print('GROUP: ' + group['id'] + ' | ' + group['name'] + ' | ' + group['siteId'])
                            dictGroups[group["id"]] = group["name"]
                    if cursor:
                        paramsnext = (
                            f"/web/api/{api_version}/groups?limit=100&cursor={cursor}&countOnly=false&tenant=true"
                        )
                        url = hostname.get() + paramsnext
                    else:
                        url = None

    async def exceptions_to_csv(querytype, session, scope, exparam):
        firstrunpath = True
        firstruncert = True
        firstrunbrowser = True
        firstrunfile = True
        firstrunhash = True

        params = (
            f"/web/api/{api_version}/exclusions?limit=1000&type={querytype}&countOnly=false"
        )
        url = hostname.get() + params + exparam
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                if response.status != 200:
                    logger.error(
                        f"Status: {str(response.status)} Problem with the request. Exiting."
                    )
                    logger.error("Details of above: " + url)
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for data in data:
                            if querytype == "path":
                                f = csv.writer(
                                    open(
                                        "exceptions_path.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrunpath:
                                    tmp = []
                                    tmp.append("Scope")
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrunpath = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "certificate":
                                f = csv.writer(
                                    open(
                                        "exceptions_certificate.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstruncert:
                                    tmp = []
                                    tmp.append("Scope")
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstruncert = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "browser":
                                f = csv.writer(
                                    open(
                                        "exceptions_browser.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrunbrowser:
                                    tmp = []
                                    tmp.append("Scope")
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrunbrowser = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "file_type":
                                f = csv.writer(
                                    open(
                                        "exceptions_file_type.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrunfile:
                                    tmp = []
                                    tmp.append("Scope")
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrunfile = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                            elif querytype == "white_hash":
                                f = csv.writer(
                                    open(
                                        "exceptions_white_hash.csv",
                                        "a+",
                                        newline="",
                                        encoding="utf-8",
                                    )
                                )
                                if firstrunhash:
                                    tmp = []
                                    tmp.append("Scope")
                                    for key, value in data.items():
                                        tmp.append(key)
                                    f.writerow(tmp)
                                    firstrunhash = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                f.writerow(tmp)

                    if cursor:
                        paramsnext = (
                            f"/web/api/{api_version}/exclusions?limit=1000&type={querytype}&countOnly=false&cursor={cursor}"
                        )
                        url = hostname.get() + paramsnext + exparam
                    else:
                        url = None

    async def run(scope):
        async with aiohttp.ClientSession() as session:

            if scope == "Account":
                exparam = "&accountIds="
                l = len(dictAccounts.items())
                i = 0
                for key, value in dictAccounts.items():
                    typepath = asyncio.create_task(
                        exceptions_to_csv(
                            "path",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typecert = asyncio.create_task(
                        exceptions_to_csv(
                            "certificate",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typebrowser = asyncio.create_task(
                        exceptions_to_csv(
                            "browser",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typefile_type = asyncio.create_task(
                        exceptions_to_csv(
                            "file_type",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typewhite_hash = asyncio.create_task(
                        exceptions_to_csv(
                            "white_hash",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    await typefile_type
                    await typebrowser
                    await typecert
                    await typepath
                    await typewhite_hash
                    i = i + 1
            elif scope == "Site":
                exparam = "&siteIds="
                l = len(dictSites.items())
                i = 0
                for key, value in dictSites.items():
                    typepath = asyncio.create_task(
                        exceptions_to_csv(
                            "path",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typecert = asyncio.create_task(
                        exceptions_to_csv(
                            "certificate",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typebrowser = asyncio.create_task(
                        exceptions_to_csv(
                            "browser",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typefile_type = asyncio.create_task(
                        exceptions_to_csv(
                            "file_type",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typewhite_hash = asyncio.create_task(
                        exceptions_to_csv(
                            "white_hash",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    await typefile_type
                    await typebrowser
                    await typecert
                    await typepath
                    await typewhite_hash
                    i = i + 1
            elif scope == "Group":
                exparam = "&groupIds="
                l = len(dictGroups.items())
                i = 0
                for key, value in dictGroups.items():
                    typepath = asyncio.create_task(
                        exceptions_to_csv(
                            "path",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typecert = asyncio.create_task(
                        exceptions_to_csv(
                            "certificate",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typebrowser = asyncio.create_task(
                        exceptions_to_csv(
                            "browser",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typefile_type = asyncio.create_task(
                        exceptions_to_csv(
                            "file_type",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    typewhite_hash = asyncio.create_task(
                        exceptions_to_csv(
                            "white_hash",
                            session,
                            scope + "|" + value + " | " + key,
                            exparam + key,
                        )
                    )
                    await typefile_type
                    await typebrowser
                    await typecert
                    await typepath
                    await typewhite_hash
                    i = i + 1
            elif scope == "Global":
                exparam = ""
                key = ""
                typepath = asyncio.create_task(
                    exceptions_to_csv("path", session, scope, exparam + key)
                )
                typecert = asyncio.create_task(
                    exceptions_to_csv("certificate", session, scope, exparam + key)
                )
                typebrowser = asyncio.create_task(
                    exceptions_to_csv("browser", session, scope, exparam + key)
                )
                typefile_type = asyncio.create_task(
                    exceptions_to_csv("file_type", session, scope, exparam + key)
                )
                typewhite_hash = asyncio.create_task(
                    exceptions_to_csv("white_hash", session, scope, exparam + key)
                )
                await typefile_type
                await typebrowser
                await typecert
                await typepath
                await typewhite_hash

    async def runAccounts():
        async with aiohttp.ClientSession() as session:

            accounts = asyncio.create_task(getAccounts(session))
            await accounts

    async def runSites():
        async with aiohttp.ClientSession() as session:

            sites = asyncio.create_task(getSites(session))
            await sites

    async def runGroups():
        async with aiohttp.ClientSession() as session:

            groups = asyncio.create_task(getGroups(session))
            await groups

    def getScope():
        r = requests.get(
            hostname.get() + f"/web/api/{api_version}/user",
            headers=headers,
            proxies={"http": proxy},
        )
        if r.status_code == 200:
            data = r.json()
            return data["data"]["scope"]
        else:
            logger.error(
                f"Status: {str(r.status_code)} Problem with the request. Details {str(r.text)}"
            )

    dictAccounts = {}
    dictSites = {}
    dictGroups = {}
    tokenscope = getScope()

    if tokenscope != "site":
        logger.info(f"Getting account/site/group structure for {hostname.get()}")
        loop = asyncio.get_event_loop()
        loop.run_until_complete(runAccounts())

    loop = asyncio.get_event_loop()
    loop.run_until_complete(runSites())

    loop = asyncio.get_event_loop()
    loop.run_until_complete(runGroups())
    logger.info("Finished getting account/site/group structure!")
    logger.info(
        f"Accounts found: {str(len(dictAccounts))} | Sites found: {str(len(dictSites))} | Groups found: {str(len(dictGroups))}"
    )

    if tokenscope == "global":
        logger.info("Getting GLOBAL scope exceptions...")
        scope = "Global"
        loop = asyncio.get_event_loop()
        loop.run_until_complete(run(scope))

    if tokenscope != "site":
        logger.info("Getting ACCOUNT scope exceptions...")
        scope = "Account"
        loop = asyncio.get_event_loop()
        loop.run_until_complete(run(scope))

    logger.info("Getting SITE scope exceptions...")
    scope = "Site"
    loop = asyncio.get_event_loop()
    loop.run_until_complete(run(scope))

    logger.info("Getting GROUP scope exceptions...")
    scope = "Group"
    loop = asyncio.get_event_loop()
    loop.run_until_complete(run(scope))

    logger.info("Creating XLSX...")

    filename = "Exceptions"
    workbook = Workbook(filename + ".xlsx")
    csvs = [
        "exceptions_path.csv",
        "exceptions_certificate.csv",
        "exceptions_browser.csv",
        "exceptions_file_type.csv",
        "exceptions_white_hash.csv",
    ]
    for csvfile in csvs:
        worksheet = workbook.add_worksheet(csvfile.split(".")[0])
        if os.path.isfile(csvfile):
            with open(csvfile, "r", encoding="utf8") as f:
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        worksheet.write(r, c, col)
        if os.path.exists(csvfile):
            os.remove(csvfile)
    workbook.close()
    logger.info(f"Done! Created the file {filename}.xlsx")


def selectCSVFile():
    file = tkinter.filedialog.askopenfilename()
    inputcsv.set(file)


# Login Menu Frame #############################
tk.Label(master=loginMenuFrame, text=gui.logo, font="TkFixedFont", fg=gui.logo_color).grid(
    row=0, column=0, columnspan=1, pady=20, padx=20
)
consoleAddressLabel = tk.Label(
    master=loginMenuFrame,
    text="Management Console URL:",
)
consoleAddressEntry = ttk.Entry(master=loginMenuFrame, width=80)

apikTokenLabel = tk.Label(
    master=loginMenuFrame,
    text="API Token:"
)
apikTokenEntry = ttk.Entry(master=loginMenuFrame, width=80)

proxyLabel = tk.Label(
    master=loginMenuFrame,
    text="Proxy (if required):",
)
proxyEntry = ttk.Entry(master=loginMenuFrame, width=80)
useSSLSwitch = ttk.Checkbutton(master=loginMenuFrame, text="Use SSL", style="Switch", variable=useSSL, onvalue=True, offvalue=False)
loginButton = ttk.Button(
    master=loginMenuFrame, text="Login", command=login
)
tk.Label(master=loginMenuFrame, text="API: {}".format(api_version)).grid(
    row=9, column=0, pady=10, sticky='s'
)
consoleAddressLabel.grid(row=1, column=0, pady=2)
consoleAddressEntry.grid(row=2, column=0, pady=2)
apikTokenLabel.grid(row=3, column=0, pady=(10,2))
apikTokenEntry.grid(row=4, column=0, pady=2)
proxyLabel.grid(row=5, column=0, pady=(10,2))
proxyEntry.grid(row=6, column=0, pady=2)
useSSLSwitch.grid(row=7, column=0, pady=10)
loginButton.grid(row=8, column=0, columnspan=2, ipady=5, pady=10)
loginMenuFrame.pack()

# Main Menu Frame #############################
tk.Label(master=mainMenuFrame, text=gui.logo, font="TkFixedFont", fg=gui.logo_color).grid(
    row=0, column=0, columnspan=2, pady=20, padx=20
)
ttk.Button(
    master=mainMenuFrame,
    text="Export Deep Visiblity Events",
    command=partial(switchFrames, exportFromDVFrame),
).grid(row=1, column=0, sticky="ew", ipady=5, pady=10, padx=(20,10))
ttk.Button(
    master=mainMenuFrame,
    text="Export and Search Activity Log",
    command=partial(switchFrames, exportActivityLogFrame),
).grid(row=1, column=1, sticky="ew", ipady=5, pady=10, padx=(10,20))
ttk.Button(
    master=mainMenuFrame,
    text="Upgrade Agents",
    command=partial(switchFrames, upgradeFromCSVFrame),
).grid(row=2, column=0, sticky="ew", ipady=5, pady=10, padx=(20,10))
ttk.Button(
    master=mainMenuFrame,
    text="Move Agents",
    command=partial(switchFrames, moveAgentsFrame),
).grid(row=2, column=1, sticky="ew", ipady=5, pady=10, padx=(10,20))
ttk.Button(
    master=mainMenuFrame,
    text="Assign Customer Identifier",
    command=partial(switchFrames, assignCustomerIdentifierFrame),
).grid(row=3, column=0, sticky="ew", ipady=5, pady=10, padx=(20,10))
ttk.Button(
    master=mainMenuFrame,
    text="Decomission Agents",
    command=partial(switchFrames, decomissionAgentsFrame),
).grid(row=3, column=1, sticky="ew", ipady=5, pady=10, padx=(10,20))
ttk.Button(
    master=mainMenuFrame,
    text="Export All Endpoints",
    command=partial(switchFrames, exportAllAgentsFrame),
).grid(row=4, column=0, sticky="ew", ipady=5, pady=10, padx=(20,10))
ttk.Button(
    master=mainMenuFrame,
    text="Export Exclusions",
    command=partial(switchFrames, exportExclusionsFrame),
).grid(row=4, column=1, sticky="ew", ipady=5, pady=10, padx=(10,20))
tk.Label(
    master=mainMenuFrame,
    text="Note: Many of the processes can take a while to run. Be patient."
).grid(row=5, column=0, columnspan=2, padx=20, pady=20, sticky='s')

# Export from DV Frame #############################
tk.Label(
    master=exportFromDVFrame,
    text="Export Deep Visiblity Events",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportFromDVFrame,
    text="Export Deep Visibility events to a CSV by query ID as reference"
).grid(row=1, column=0, padx=20, pady=10)
tk.Label(
    master=exportFromDVFrame,
    text="1. Input Deep Visibility Query ID"
).grid(row=2, column=0, pady=2)
queryIdEntry = ttk.Entry(master=exportFromDVFrame, width=80)
queryIdEntry.grid(row=3, column=0, pady=10)
ttk.Button(
    master=exportFromDVFrame,
    text="Export",
    command=exportFromDV,
).grid(row=4, column=0, pady=10)
ttk.Button(
    master=exportFromDVFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=5, column=0, ipadx=10, pady=10)


# Export and Search Activity Log Frame #############################
tk.Label(
    master=exportActivityLogFrame,
    text="Export and Search Activity Log",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(master=exportActivityLogFrame, text="Search Management Console Activity log and export results.").grid(
    row=1, column=0, padx=20, pady=10
)
tk.Label(master=exportActivityLogFrame, text="1. Input FROM date (yyyy-MM-dd)").grid(
    row=2, column=0, pady=2
)
dateFrom = fromDateEntry = ttk.Entry(master=exportActivityLogFrame, width=40)
fromDateEntry.grid(row=3, column=0, pady=10)
tk.Label(master=exportActivityLogFrame, text="2. Input TO date (yyyy-MM-dd)").grid(
    row=4, column=0, pady=2
)
dateTo = toDateEntry = ttk.Entry(master=exportActivityLogFrame, width=40)
toDateEntry.grid(row=5, column=0, pady=10)
tk.Label(
    master=exportActivityLogFrame, text="3. Input search string"
).grid(row=6, column=0, pady=2)
stringSearchEntry = ttk.Entry(master=exportActivityLogFrame, width=80)
stringSearchEntry.grid(row=7, column=0, pady=2)
ttk.Button(
    master=exportActivityLogFrame,
    text="Search",
    command=partial(exportActivityLog, True),
).grid(row=8, column=0, pady=10)
ttk.Button(
    master=exportActivityLogFrame,
    text="Export",
    command=partial(exportActivityLog, False),
).grid(row=9, column=0, pady=10)
ttk.Button(
    master=exportActivityLogFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=10, column=0, ipadx=10, pady=10)


# Upgrade Agents Frame #############################
tk.Label(
    master=upgradeFromCSVFrame,
    text="Upgrade Agents",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=upgradeFromCSVFrame, text="Upgrade Agents listed in a CSV to a specific Package (referenced by ID)"
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=upgradeFromCSVFrame, text="1. Export Packages List to source Package ID"
).grid(row=2, column=0, padx=20, pady=2)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Export Packages List",
    command=partial(upgradeFromCSV, True),
).grid(row=3, column=0, pady=10)
tk.Label(
    master=upgradeFromCSVFrame, text="2. Insert the Package ID"
).grid(row=4, column=0, pady=2)
packageIDEntry = ttk.Entry(master=upgradeFromCSVFrame, width=80)
packageIDEntry.grid(row=5, column=0, pady=2)
tk.Label(
    master=upgradeFromCSVFrame,
    text="3. Select a CSV file containing a single column of endpoint names to upgrade"
).grid(row=6, column=0, padx=20, pady=2)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=7, column=0, pady=2)
tk.Label(master=upgradeFromCSVFrame, textvariable=inputcsv).grid(
    row=8, column=0, pady=2
)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Submit",
    command=partial(upgradeFromCSV, False),
).grid(row=9, column=0, pady=10)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=10, column=0, ipadx=10, pady=10)


# Move Agents Frame #############################
tk.Label(
    master=moveAgentsFrame,
    text="Move Agents",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=moveAgentsFrame,
    text="1. Export Groups List to get group IDs"
).grid(row=1, column=0, pady=2)
ttk.Button(
    master=moveAgentsFrame,
    text="Export Groups List",
    command=partial(moveAgents, True),
).grid(row=2, column=0, pady=10)
tk.Label(
    master=moveAgentsFrame,
    text="2. Select a CSV file constructed of three columns:\nendpoints names, target group IDs, target site IDs"
).grid(row=3, column=0, padx=20, pady=10)
ttk.Button(
    master=moveAgentsFrame, text="Browse", command=selectCSVFile
).grid(row=4, column=0, pady=10)
tk.Label(master=moveAgentsFrame, textvariable=inputcsv).grid(row=5, column=0, pady=10)
ttk.Button(
    master=moveAgentsFrame,
    text="Submit",
    command=partial(moveAgents, False),
).grid(row=6, column=0, pady=10)
ttk.Button(
    master=moveAgentsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=7, column=0, ipadx=10, pady=10)


# Assign Customer Identifier Frame #############################
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="Assign Customer Identifier",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="Assign a Customer Identifier to one or more Agents.\n\n1. Input the Customer Identifier to assign."
).grid(row=1, column=0, pady=2)
customerIdentifierEntry = ttk.Entry(master=assignCustomerIdentifierFrame, width=80)
customerIdentifierEntry.grid(row=2, column=0, pady=(2,10))
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="2. Select a CSV file containing a single column with endpoint names"
).grid(row=3, column=0, padx=20, pady=2)
ttk.Button(
    master=assignCustomerIdentifierFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=4, column=0, pady=10)
tk.Label(master=assignCustomerIdentifierFrame, textvariable=inputcsv).grid(
    row=5, column=0, pady=10
)
ttk.Button(
    master=assignCustomerIdentifierFrame,
    text="Submit",
    command=assignCustomerIdentifier,
).grid(row=6, column=0, pady=10)
ttk.Button(
    master=assignCustomerIdentifierFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=7, column=0, ipadx=10, pady=10)


# Decomission Agents from CSV Frame #############################
tk.Label(
    master=decomissionAgentsFrame,
    text="Decomission Agents",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=decomissionAgentsFrame,
    text="1. Select a CSV file containing a single column of endpoint names to be decomissioned"
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=decomissionAgentsFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=2, column=0, pady=10)
tk.Label(master=decomissionAgentsFrame, textvariable=inputcsv).grid(
    row=3, column=0, pady=10
)
ttk.Button(
    master=decomissionAgentsFrame,
    text="Submit",
    command=decomissionAgents,
).grid(row=4, column=0, pady=10)
ttk.Button(
    master=decomissionAgentsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=5, column=0, ipadx=10, pady=10)


# Export all agents Frame #############################
tk.Label(
    master=exportAllAgentsFrame,
    text="Export Endpoint Details to CSV",
    font=("Courier", 24),
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportAllAgentsFrame, text="Exports All Agent Details to a CSV"
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=exportAllAgentsFrame,
    text="Export",
    command=exportAllAgents,
).grid(row=2, column=0, pady=10)
ttk.Button(
    master=exportAllAgentsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=3, column=0, ipadx=10, pady=10)


# Export Exclusions #############################
tk.Label(
    master=exportExclusionsFrame, text="Export Exclusions to CSV", font=("Courier", 24)
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportExclusionsFrame, text="Exports All Exclusions to a CSV"
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=exportExclusionsFrame,
    text="Export",
    command=exportExclusions,
).grid(row=2, column=0, pady=10)
ttk.Button(
    master=exportExclusionsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=3, column=0, ipadx=10, pady=10)


def main():
    window.mainloop()


if __name__ == "__main__":
    main()
