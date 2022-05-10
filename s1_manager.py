""" s1_manager.py 
    Source: https://github.com/DylanCS1/s1_manager
    License: MIT license - https://github.com/DylanCS1/s1_manager/blob/main/LICENSE.txt
"""

import asyncio
import csv
import datetime
import json
import logging
import os
import platform
import sys
import time
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext as ScrolledText
from functools import partial
from tkinter import UNDERLINE, ttk

import aiohttp
import requests
from PIL import Image, ImageTk
from xlsxwriter.workbook import Workbook

# CONSTS
__version__ = "2022.1.1"
api_version = "v2.1"
dir_path = os.path.dirname(os.path.realpath(__file__))
query_limits = "limit=1000"

# LOG SETTINGS
if len(sys.argv) > 1 and sys.argv[1] == "--debug":
    LOG_LEVEL = logging.DEBUG
    LOG_NAME = f"s1_manager_debug_{datetime.datetime.now().strftime('%Y-%m-%d')}.log"
else:
    LOG_LEVEL = logging.INFO
    LOG_NAME = f"s1_manager_{datetime.datetime.now().strftime('%Y-%m-%d')}.log"
if LOG_LEVEL == logging.DEBUG:
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)s - %(message)s"
else:
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

# WINDOW SETTINGS
window = tk.Tk()
window.title("S1 Manager")
window.iconbitmap(os.path.join(dir_path, ".ICO/s1_manager.ico"))
window.minsize(850, 650)

# THEME
window.tk.call(
    "source", os.path.join(dir_path, ".THEME/forest-dark.tcl")
)  # https://github.com/rdbende/Forest-ttk-theme
logo = ImageTk.PhotoImage(Image.open(os.path.join(dir_path, ".ICO/s1_manager.png")))
ttk.Style().theme_use("forest-dark")
frame_title_font = ("Courier", 24, UNDERLINE)
frame_subtitle_font_underline = ("Arial", 14, UNDERLINE)
frame_subtitle_font = ("Arial", 12)
frame_subnote_font = ("Arial", 10)
frame_note_fg_color = "red"
st_font = "TkFixedFont"

# FRAME CONSTS
loginMenuFrame = ttk.Frame()
mainMenuFrame = ttk.Frame()
exportFromDVFrame = ttk.Frame()
exportActivityLogFrame = ttk.Frame()
exportEndpointsFrame = ttk.Frame()
exportEndpointTagsFrame = ttk.Frame()
exportExclusionsFrame = ttk.Frame()
exportLocalConfigFrame = ttk.Frame()
exportUsersFrame = ttk.Frame()
exportRangerInvFrame = ttk.Frame()
upgradeFromCSVFrame = ttk.Frame()
moveAgentsFrame = ttk.Frame()
assignCustomerIdentifierFrame = ttk.Frame()
decommissionAgentsFrame = ttk.Frame()
manageEndpointTagsFrame = ttk.Frame()
error = tk.StringVar()
hostname = tk.StringVar()
apitoken = tk.StringVar()
proxy = tk.StringVar()
inputcsv = tk.StringVar()
useSSL = tk.BooleanVar()
useSSL.set(True)
useSchedule = tk.BooleanVar()
useSchedule.set(False)


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
    """Function to test login using APIToken or Token"""

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
    else:
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
        r.raise_for_status()
        if r.status_code == 200:
            return headers, True
        else:
            return 0, False


def login():
    """Function to handle login actions"""
    hostname.set(consoleAddressEntry.get())
    apitoken.set(apikTokenEntry.get())
    proxy.set(proxyEntry.get())
    global headers

    if not hostname.get() or not apitoken.get():
        tk.Label(
            master=loginMenuFrame,
            text="'Management Console URL' and 'API Token' cannot be empty.",
            fg=frame_note_fg_color,
            font=frame_subnote_font,
        ).grid(row=11, column=0, columnspan=2, pady=10)
    else:
        headers, login_succ = testLogin(hostname.get(), apitoken.get(), proxy.get())
        if login_succ:
            loginMenuFrame.pack_forget()
            mainMenuFrame.pack()
        else:
            tk.Label(
                master=loginMenuFrame,
                text="Login to the management console failed. Please check your credentials and try again",
                fg=frame_note_fg_color,
                font=frame_subnote_font,
            ).grid(row=11, column=0, columnspan=2, pady=10)


def goBacktoMainPage():
    """Function to handle moving back to the Main Menu Frame"""
    _list = window.winfo_children()
    for item in _list:
        if item.winfo_children():
            _list.extend(item.winfo_children())
    for item in _list:
        if isinstance(item, tk.Toplevel) is not True:
            item.pack_forget()
    mainMenuFrame.pack()


def switchFrames(framename):
    """Function to handle switching tkinter frames"""
    inputcsv.set("")
    mainMenuFrame.pack_forget()
    framename.pack()


def exportFromDV():
    """Function to export events from Deep Visibility by DV query ID"""
    st = ScrolledText.ScrolledText(
        master=exportFromDVFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    # Filename variables
    dv_file = "dv_file.csv"
    dv_ip = "dv_ip.csv"
    dv_url = "dv_url.csv"
    dv_dns = "dv_dns.csv"
    dv_process = "dv_process.csv"
    dv_registry = "dv_registry.csv"
    dv_scheduled_task = "dv_scheduled_task.csv"

    async def dv_query_to_csv(
        querytype, session, hostname, dv_query_id, headers, firstrun, proxy
    ):
        params = f"/web/api/{api_version}/dv/events/{querytype}?queryId={dv_query_id}"
        url = hostname + params
        while url:
            async with session.get(
                url, headers=headers, proxy=proxy, ssl=useSSL.get()
            ) as response:
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy,
                    useSSL.get(),
                )
                if response.status != 200:
                    logger.error(
                        "HTTP Response Code: %d %s - There was a problem with the request to %s.",
                        response.status,
                        response.reason,
                        url,
                    )
                    break
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for data in data:
                            logging.debug("Query type is %s", querytype)
                            if querytype == "file":
                                f = csv.writer(
                                    open(
                                        dv_file,
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
                                    open(dv_ip, "a+", newline="", encoding="utf-8")
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
                                    open(dv_url, "a+", newline="", encoding="utf-8")
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
                                    open(dv_dns, "a+", newline="", encoding="utf-8")
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
                                        dv_process,
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
                                        dv_registry,
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
                                        dv_scheduled_task,
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
                        paramsnext = f"/web/api/{api_version}/dv/events/{querytype}?cursor={cursor}&queryId={dv_query_id}&{query_limits}"
                        url = hostname + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
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
        logger.info("Processing DV Query ID: %s", dv_query_id)
        dv_query_id = dv_query_id.split(",")
        if platform.system() == "Windows":
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        asyncio.run(run(hostname.get(), dv_query_id, apitoken.get(), proxy.get()))
        xlsx_filename = "-"
        xlsx_filename = f"DV_Export_{xlsx_filename.join(dv_query_id)}.xlsx"
        workbook = Workbook(xlsx_filename)
        csvs = [
            dv_file,
            dv_ip,
            dv_url,
            dv_dns,
            dv_process,
            dv_registry,
            dv_scheduled_task,
        ]
        for csvfile in csvs:
            worksheet = workbook.add_worksheet(csvfile.split(".")[0])
            if os.path.isfile(csvfile):
                with open(csvfile, "r", encoding="utf8") as f:
                    logger.debug("Reading %s and writing to %s", csvfile, workbook)
                    reader = csv.reader(f)
                    for r, row in enumerate(reader):
                        for c, col in enumerate(row):
                            worksheet.write(r, c, col)
                logger.debug("Deleting %s", csvfile)
                os.remove(csvfile)
        workbook.close()
        logger.info("Done! Created the file %s\n", xlsx_filename)
    else:
        logger.error("Please enter a valid DV Query ID and try again.", dv_query_id)


def exportActivityLog(searchOnly):
    """Function to search for Activity events by date range or export Activity events"""
    st = ScrolledText.ScrolledText(
        master=exportActivityLogFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    os.environ["TZ"] = "UTC"
    p = "%Y-%m-%d"
    fromdate_epoch = str(int(time.mktime(time.strptime(dateFrom.get(), p)))) + "000"
    todate_epoch = str(int(time.mktime(time.strptime(dateTo.get(), p)))) + "000"
    logger.debug("Input FROM Date: %s Input TO Date: %s", dateFrom.get(), dateTo.get())
    logger.debug(
        "Epoch-converted FROM Date: %s Epoch-converted TO Date: %s",
        fromdate_epoch,
        todate_epoch,
    )
    if dateFrom.get() and dateTo.get():
        url = (
            hostname.get()
            + f"/web/api/{api_version}/activities?{query_limits}&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&includeHidden=false"
        )
        logger.debug("Search only state: %s", searchOnly)
        if searchOnly:
            logger.info("Starting search for '%s'", stringSearchEntry.get())
            while url:
                response = requests.get(
                    url,
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Status: %s Problem with the request. Details - %s",
                        str(response.status_code),
                        str(response.text),
                    )
                    break
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
                                    "%s - %s - %s",
                                    item["createdAt"],
                                    item["primaryDescription"],
                                    item["secondaryDescription"],
                                )
                            elif item["secondaryDescription"]:
                                if (
                                    stringSearchEntry.get().upper()
                                    in item["secondaryDescription"].upper()
                                ):
                                    logger.info(
                                        "%s - %s - %s",
                                        item["createdAt"],
                                        item["primaryDescription"],
                                        item["secondaryDescription"],
                                    )
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/activities?{query_limits}&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&cursor={cursor}&includeHidden=false"
                        url = hostname.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None
        else:
            datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
            csv_filename = f"Activity_Log_Export_{datestamp}.csv"
            logger.info("Creating and opening %s", csv_filename)
            f = csv.writer(
                open(
                    csv_filename,
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
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Status: %s Problem with the request. Details - %s",
                        str(response.status_code),
                        str(response.text),
                    )
                    break
                else:
                    data = response.json()
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        if firstrun:
                            tmp = []
                            for key, value in data[0].items():
                                tmp.append(key)
                            logger.debug("Writing column headers on first run.")
                            f.writerow(tmp)
                            logger.debug(
                                "First run through the data set complete, setting firstrun to False"
                            )
                            firstrun = False
                        for item in data:
                            tmp = []
                            for key, value in item.items():
                                tmp.append(value)
                            logger.debug(
                                "Writing entry to CSV: %s - %s - %s",
                                item["createdAt"],
                                item["primaryDescription"],
                                item["secondaryDescription"],
                            )
                            f.writerow(tmp)
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/activities?{query_limits}&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&cursor={cursor}&includeHidden=false"
                        url = hostname.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None
            logger.info("Done! Output file is - %s\n", csv_filename)
    else:
        logger.error("You must state a FROM date and a TO date")


def upgradeFromCSV(justPackages):
    """Function to upgrade Agents via API"""
    st = ScrolledText.ScrolledText(
        master=upgradeFromCSVFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    csv_filename = f"Available_Packages_List_{datestamp}.csv"

    logger.debug("Just packages set to: %s", justPackages)
    if justPackages:
        params = f"/web/api/{api_version}/update/agent/packages?sortBy=updatedAt&sortOrder=desc&countOnly=false&{query_limits}"
        url = hostname.get() + params
        f = csv.writer(open(csv_filename, "a+", newline="", encoding="utf-8"))
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
            logger.debug(
                "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                proxy.get(),
                useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Status: %s Problem with the request. Details - %s",
                    str(response.status_code),
                    str(response.text),
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
                    paramsnext = f"/web/api/{api_version}/update/agent/packages?sortBy=updatedAt&sortOrder=desc&{query_limits}&cursor={cursor}&countOnly=false"
                    url = hostname.get() + paramsnext
                    logger.debug("Found next cursor: %s", cursor)
                else:
                    logger.debug("No cursor found, setting URL to None")
                    url = None
        logger.info("SentinelOne agent packages list written to: %s", csv_filename)
    else:
        with open(inputcsv.get()) as csv_file:
            logger.debug("Reading CSV: %s", inputcsv.get())
            csv_reader = csv.reader(csv_file, delimiter=",")
            line_count = 0
            logger.debug("Use Schedule value: %s", useSchedule.get())
            for row in csv_reader:
                logger.info("Upgrading endpoint named - %s", row[0])
                url = (
                    hostname.get()
                    + f"/web/api/{api_version}/agents/actions/update-software"
                )
                body = {
                    "filter": {"computerName": row[0]},
                    "data": {
                        "packageId": packageIDEntry.get(),
                        "isScheduled": useSchedule.get(),
                    },
                }
                response = requests.post(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tData: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    json.dumps(body),
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Failed to upgrade endpoint %s Error code: %s Description: %s",
                        row[0],
                        str(response.status_code),
                        str(response.text),
                    )
                else:
                    data = response.json()
                    logger.info(
                        "Sent upgrade command to %s endpoints", data["data"]["affected"]
                    )
                    if useSchedule.get():
                        logger.info(
                            "Upgrade should follow schedule defined in Management Console."
                        )
                line_count += 1
            if line_count < 1:
                logger.info("Finished! Input file %s was empty.", inputcsv.get())
            else:
                logger.info("Finished! Processed %d lines.", line_count)


def moveAgents(justGroups):
    """Function to move Agents using API"""
    st = ScrolledText.ScrolledText(master=moveAgentsFrame, state="disabled", height=10)
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    if justGroups:
        logger.debug("Just move agents to groups: %s", justGroups)
        params = f"/web/api/{api_version}/groups?isDefault=false&{query_limits}&type=static&countOnly=false"
        url = hostname.get() + params
        csv_filename = "Group_To_ID_Map.csv"
        f = csv.writer(open(csv_filename, "a+", newline="", encoding="utf-8"))
        f.writerow(["Name", "ID", "Site ID", "Created By"])
        while url:
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            logger.debug(
                "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                proxy.get(),
                useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Status: %s Problem with the request. Details - %s",
                    str(response.status_code),
                    str(response.text),
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
                    paramsnext = f"/web/api/{api_version}/groups?isDefault=false&{query_limits}&type=static&cursor={cursor}&countOnly=false"
                    url = hostname.get() + paramsnext
                    logger.debug("Found next cursor: %s", cursor)
                else:
                    logger.debug("No cursor found, setting URL to None")
                    url = None
        logger.info("Added group mapping to the file %s", csv_filename)
    else:
        with open(inputcsv.get()) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            line_count = 0
            for row in csv_reader:
                logger.info("Moving endpoint name %s to Site ID %s", row[0], row[2])
                url = (
                    hostname.get()
                    + f"/web/api/{api_version}/agents/actions/move-to-site"
                )
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
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tData: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    json.dumps(body),
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Failed to transfer endpoint %s to site %s Error code: %s Description: %s",
                        row[0],
                        row[2],
                        str(response.status_code),
                        str(response.text),
                    )
                else:
                    data = response.json()
                    logger.info("Moved %s endpoints", data["data"]["affected"])
                logger.info("Moving endpoint name %s to Group ID %s", row[0], row[1])
                url = (
                    hostname.get()
                    + f"/web/api/{api_version}/groups/"
                    + row[1]
                    + "/move-agents"
                )
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
                        "Failed to transfer endpoint %s to group %s Error code: %s Description: %s",
                        row[0],
                        row[1],
                        str(response.status_code),
                        str(response.text),
                    )
                else:
                    data = response.json()
                    logger.info("Moved %s endpoints", data["data"]["agentsMoved"])
                line_count += 1
            if line_count < 1:
                logger.info("Finished! Input file %s was empty.", inputcsv.get())
            else:
                logger.info("Finished! Processed %d lines.", line_count)


def assignCustomerIdentifier():
    """Function to add a Customer Identifier to one or more Agents via API"""
    st = ScrolledText.ScrolledText(
        master=assignCustomerIdentifierFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    with open(inputcsv.get()) as csv_file:
        logger.debug("Reading CSV: %s", inputcsv.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0
        for row in csv_reader:
            logger.info("Updating customer identifier for endpoint - %s", row[0])
            url = (
                hostname.get()
                + f"/web/api/{api_version}/agents/actions/set-external-id"
            )
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
            logger.debug(
                "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                proxy.get(),
                useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to update customer identifier for endpoint %s Error code: %s Description: %s",
                    row[0],
                    str(response.status_code),
                    str(response.text),
                )
            else:
                r = response.json()
                affected_num_of_endpoints = r["data"]["affected"]
                if affected_num_of_endpoints < 1:
                    logger.info("No endpoint matched the name %s", row[0])
                elif affected_num_of_endpoints > 1:
                    logger.info(
                        "%s endpoints matched the name %s , customer identifier was updated for all",
                        affected_num_of_endpoints,
                        row[0],
                    )
                else:
                    logger.info("Successfully updated the customer identifier")
            line_count += 1
        if line_count < 1:
            logger.info("Finished! Input file %s was empty.", inputcsv.get())
        else:
            logger.info("Finished! Processed %d lines.", line_count)


def exportAllAgents():
    """Function to export a list of all Agents and details to CSV"""
    st = ScrolledText.ScrolledText(
        master=exportEndpointsFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    logger.debug("User selected %s file type", endpointOutputType.get())
    output_file_name = f"Export_Endpoints_{datestamp}"
    csv_file = output_file_name + ".csv"
    xlsx_file = output_file_name + ".xlsx"

    firstrun = True
    url = (
        hostname.get()
        + f"/web/api/{api_version}/agents?{query_limits}&sortBy=computerName&sortOrder=asc"
    )

    logger.info("Starting to request endpoint data.")
    while url:
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": proxy.get(), "https": proxy.get()},
            verify=useSSL.get(),
        )
        logger.debug(
            "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            proxy.get(),
            useSSL.get(),
        )
        if response.status_code != 200:
            logger.error(
                "Status: %s Problem with the request. Details - %s ",
                str(response.status_code),
                str(response.text),
            )
            break
        else:
            data = response.json()
            cursor = data["pagination"]["nextCursor"]
            total_endpoints = data["pagination"]["totalItems"]
            data = data["data"]

            logger.debug(
                "%s total endpoints in data. Opening %s to start writing",
                total_endpoints,
                csv_file,
            )
            f = csv.writer(open(csv_file, "a+", newline="", encoding="utf-8"))
            if firstrun:
                tmp = []
                for key, value in data[0].items():
                    tmp.append(key)
                logger.debug(
                    "Writing column headers to first row: %s",
                    tmp,
                )
                f.writerow(tmp)
                logger.debug("First run complete, setting firstrun to False")
                firstrun = False
            for item in data:
                tmp = []
                for key, value in item.items():
                    tmp.append(value)
                logger.debug(
                    "Writing data to new row: %s",
                    tmp,
                )
                f.writerow(tmp)

            if cursor:
                paramsnext = f"/web/api/{api_version}/agents?{query_limits}&sortBy=computerName&sortOrder=asc&cursor={cursor}"
                url = hostname.get() + paramsnext
                logger.debug("Next cursor found, updating URL: %s", url)
            else:
                logger.debug("No cursor found, setting URL to None")
                url = None

    if endpointOutputType.get() == "xlsx":
        logger.info("Creating new XLSX: %s", xlsx_file)
        workbook = Workbook(xlsx_file)
        logger.debug("Adding new worksheet: 'Endpoints'")
        worksheet = workbook.add_worksheet("Endpoints")
        if os.path.isfile(csv_file):
            with open(csv_file, "r", encoding="utf8") as f:
                logger.info("Reading %s and writing to %s", csv_file, xlsx_file)
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        worksheet.write(r, c, col)
            logger.debug("Deleting %s", csv_file)
            os.remove(csv_file)
        else:
            logger.error("%s not found.", csv_file)
        logger.debug("Closing XLSX")
        workbook.close()

    logger.info(
        "Done! Output file is - %s.%s\n", output_file_name, endpointOutputType.get()
    )


def decommissionAgents():
    """Function to decommission specified agents via API"""
    st = ScrolledText.ScrolledText(
        master=decommissionAgentsFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    with open(inputcsv.get()) as csv_file:
        logger.debug("Reading CSV: %s", inputcsv.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0
        for row in csv_reader:
            logger.info("Decommissioning Endpoint - %s", row[0])
            logger.info("Getting endpoint ID for %s", row[0])
            url = (
                hostname.get()
                + f"/web/api/{api_version}/agents?countOnly=false&computerName={row[0]}&{query_limits}"
            )
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            logger.debug(
                "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                proxy.get(),
                useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to get ID for endpoint %s Error code: %s Description: %s",
                    row[0],
                    str(response.status_code),
                    str(response.text),
                )
            else:
                r = response.json()
                totalitems = r["pagination"]["totalItems"]
                logger.debug("Total items returned: %s", totalitems)
                if totalitems < 1:
                    logger.info(
                        "Could not locate any IDs for endpoint named %s - Please note the query is CaSe SenSiTiVe",
                        row[0],
                    )
                else:
                    r = r["data"]
                    uuidslist = []
                    for item in r:
                        uuidslist.append(item["id"])
                        logger.info(
                            "Found ID %s! Adding it to be decommissioned", item["id"]
                        )
                    url = (
                        hostname.get()
                        + f"/web/api/{api_version}/agents/actions/decommission"
                    )
                    body = {"filter": {"ids": uuidslist}}
                    response = requests.post(
                        url,
                        data=json.dumps(body),
                        headers=headers,
                        proxies={"http": proxy.get(), "https": proxy.get()},
                        verify=useSSL.get(),
                    )
                    logger.debug(
                        "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                        url,
                        headers,
                        proxy.get(),
                        useSSL.get(),
                    )
                    if response.status_code != 200:
                        logger.error(
                            "Failed to decommission endpoint %s Error code: %s Description: %s",
                            row[0],
                            str(response.status_code),
                            str(response.text),
                        )
                    else:
                        r = response.json()
                        affected_num_of_endpoints = r["data"]["affected"]
                        if affected_num_of_endpoints < 1:
                            logger.info("No endpoint matched the name %s", row[0])
                        elif affected_num_of_endpoints > 1:
                            logger.info(
                                "%s endpoints matched the name %s, all of them got decommissioned",
                                affected_num_of_endpoints,
                                row[0],
                            )
                        else:
                            logger.info("Successfully decommissioned the endpoint")
            line_count += 1
        if line_count < 1:
            logger.info("Finished! Input file %s was empty.", inputcsv.get())
        else:
            logger.info("Finished! Processed %d lines.", line_count)


def exportExclusions():
    """Function to export Exclusions to CSV"""
    st = ScrolledText.ScrolledText(
        master=exportExclusionsFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    async def getAccounts(session):
        logger.info("Getting accounts data")
        params = (
            f"/web/api/{api_version}/accounts?{query_limits}"
            + "&countOnly=false&tenant=true"
        )
        url = hostname.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status != 200:
                    logger.error(
                        "HTTP Response Code: %d %s - There was a problem with the request to %s.",
                        response.status,
                        response.reason,
                        url,
                    )
                    break
                else:
                    logger.debug("Request successful. Status: %s", str(response.status))
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for account in data:
                            dictAccounts[account["id"]] = account["name"]
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/accounts?{query_limits}&cursor={cursor}&countOnly=false&tenant=true"
                        url = hostname.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def getSites(session):
        logger.info("Getting sites data")
        params = (
            f"/web/api/{api_version}/sites?{query_limits}&countOnly=false&tenant=true"
        )
        url = hostname.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status != 200:
                    logger.error(
                        "HTTP Response Code: %d %s - There was a problem with the request to %s.",
                        response.status,
                        response.reason,
                        url,
                    )
                    break
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for site in data["sites"]:
                            dictSites[site["id"]] = site["name"]
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/sites?{query_limits}&cursor={cursor}&countOnly=false&tenant=true"
                        url = hostname.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def getGroups(session):
        logger.info("Getting groups data")
        params = (
            f"/web/api/{api_version}/groups?{query_limits}&countOnly=false&tenant=true"
        )
        url = hostname.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status != 200:
                    logger.error(
                        "HTTP Response Code: %d %s - There was a problem with the request to %s.",
                        response.status,
                        response.reason,
                        url,
                    )
                    break
                else:
                    data = await (response.json())
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if data:
                        for group in data:
                            dictGroups[group["id"]] = group["name"]
                    if cursor:
                        paramsnext = f"/web/api/{api_version}/groups?{query_limits}&cursor={cursor}&countOnly=false&tenant=true"
                        url = hostname.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def exceptions_to_csv(querytype, session, scope, exparam):
        firstrunpath = True
        firstruncert = True
        firstrunbrowser = True
        firstrunfile = True
        firstrunhash = True
        logger.debug("Getting exceptions and writing to CSV")
        params = f"/web/api/{api_version}/exclusions?{query_limits}&type={querytype}&countOnly=false"
        url = hostname.get() + params + exparam
        while url:
            async with session.get(url, headers=headers, proxy=proxy.get()) as response:
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                if response.status != 200:
                    logger.error(
                        "HTTP Response Code: %d %s - There was a problem with the request to %s.",
                        response.status,
                        response.reason,
                        url,
                    )
                    break
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
                        paramsnext = f"/web/api/{api_version}/exclusions?{query_limits}&type={querytype}&countOnly=false&cursor={cursor}"
                        url = hostname.get() + paramsnext + exparam
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def run(scope):
        async with aiohttp.ClientSession() as session:

            logger.debug("Scope is: %s", scope)
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
            logger.debug("Running through accounts")
            accounts = asyncio.create_task(getAccounts(session))
            await accounts

    async def runSites():
        async with aiohttp.ClientSession() as session:
            logger.debug("Running through sites")
            sites = asyncio.create_task(getSites(session))
            await sites

    async def runGroups():
        async with aiohttp.ClientSession() as session:
            logger.debug("Running through groups")
            groups = asyncio.create_task(getGroups(session))
            await groups

    def getScope():
        logger.info("Getting user scope access")
        url = hostname.get() + f"/web/api/{api_version}/user"
        r = requests.get(
            url,
            headers=headers,
            proxies={"http": proxy},
        )
        logger.debug(
            "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            proxy.get(),
            useSSL.get(),
        )
        if r.status_code == 200:
            data = r.json()
            return data["data"]["scope"]
        else:
            logger.error(
                "Status: %s Problem with the request. Details - %s",
                str(r.status_code),
                str(r.text),
            )

    dictAccounts = {}
    dictSites = {}
    dictGroups = {}
    tokenscope = getScope()

    if tokenscope != "site":
        logger.info("Getting account/site/group structure for %s", hostname.get())
        loop = asyncio.get_event_loop()
        loop.run_until_complete(runAccounts())

    loop = asyncio.get_event_loop()
    loop.run_until_complete(runSites())

    loop = asyncio.get_event_loop()
    loop.run_until_complete(runGroups())
    logger.info("Finished getting account/site/group structure!")
    logger.info(
        "Accounts found: %s | Sites found: %s | Groups found: %s",
        str(len(dictAccounts)),
        str(len(dictSites)),
        str(len(dictGroups)),
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

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    xlsx_filename = f"Exceptions_Export_{datestamp}.xlsx"
    workbook = Workbook(xlsx_filename)
    csvs = [
        "exceptions_path.csv",
        "exceptions_certificate.csv",
        "exceptions_browser.csv",
        "exceptions_file_type.csv",
        "exceptions_white_hash.csv",
    ]
    for csvfile in csvs:
        logger.debug("Writing CSV: %s", csvfile)
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
    logger.info("Done! Created the file %s\n", xlsx_filename)


def exportEndpointTags():
    """Function to export Endpoint Tags from Console"""
    st = ScrolledText.ScrolledText(
        master=exportEndpointTagsFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    export_csv = f"Endpoint_Tags_Export_{datestamp}.csv"
    f = csv.writer(open(export_csv, "a+", newline="", encoding="utf-8"))
    firstrun = True
    url = (
        hostname.get()
        + f"/web/api/{api_version}/agents/tags?includeChildren=true&includeParents=true&{query_limits}"
    )
    while url:
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": proxy.get(), "https": proxy.get()},
            verify=useSSL.get(),
        )
        logger.debug(
            "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            proxy.get(),
            useSSL.get(),
        )
        if response.status_code != 200:
            logger.error(
                "Status: %s Problem with the request. Details - %s ",
                str(response.status_code),
                str(response.text),
            )
            break
        else:
            data = response.json()
            cursor = data["pagination"]["nextCursor"]
            data = data["data"]
            logger.info("Writing endpoint tags data to %s", export_csv)
            if data:
                if firstrun:
                    logger.debug("First run through data")
                    tmp = []
                    for key, value in data[0].items():
                        tmp.append(key)
                    logger.debug(
                        "Writing column headers to first row: %s",
                        tmp,
                    )
                    f.writerow(tmp)
                    logger.debug("First run complete, setting firstrun to False")
                    firstrun = False
                for item in data:
                    tmp = []
                    for key, value in item.items():
                        tmp.append(value)
                    f.writerow(tmp)
            if cursor:
                paramsnext = f"/web/api/{api_version}/agents/tags?includeChildren=true&includeParents=true&{query_limits}&cursor={cursor}"
                url = hostname.get() + paramsnext
                logger.debug("Next cursor found, updating URL: %s", url)
            else:
                logger.debug("No cursor found, setting URL to None")
                url = None
    logger.info("Done! Output file is - %s\n", export_csv)


def manageEndpointTags():
    """Add or Remove Endpoint Tags from Agents"""
    st = ScrolledText.ScrolledText(
        master=manageEndpointTagsFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    id_type = "computerName"
    if agentIDType.get() == "uuid":
        id_type = "uuid"

    logger.debug("Specified an ID type of: %s", id_type)

    with open(inputcsv.get()) as csv_file:
        logger.debug("Reading CSV: %s", inputcsv.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0

        for row in csv_reader:
            logger.info("Updating Endpoint Tags for %s", row[0])
            url = hostname.get() + f"/web/api/{api_version}/agents/actions/manage-tags"
            body = {
                "filter": {id_type: row[0]},
                "data": [
                    {"operation": endpointTagsAction.get(), "tagId": tagIDEntry.get()}
                ],
            }
            logger.debug(
                "Making API Call with the following:\nURL: %s\tData: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                json.dumps(body),
                headers,
                proxy.get(),
                useSSL.get(),
            )
            response = requests.post(
                url,
                data=json.dumps(body),
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to update Endpoint Tag for agent %s Error code: %s Description: %s",
                    row[0],
                    str(response.status_code),
                    str(response.text).strip(),
                )
            else:
                r = response.json()
                affected_num_of_endpoints = r["data"]["affected"]
                if affected_num_of_endpoints < 1:
                    logger.info("Endpoint Tag not updated for agent %s", row[0])
                else:
                    logger.info("Successfully updated the Endpoint Tag")
            line_count += 1
        if line_count < 1:
            logger.info("Finished! Input file %s was empty.", inputcsv.get())
        else:
            logger.info("Finished! Processed %d lines.", line_count)


def exportLocalConfig():
    """Export Agent Local Config"""
    st = ScrolledText.ScrolledText(
        master=exportLocalConfigFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    json_file = f"Local_Config_Export_{datestamp}.json"

    with open(inputcsv.get()) as csv_file:
        logger.debug("Reading CSV: %s", inputcsv.get())
        csv_reader = csv.reader(csv_file, delimiter=",")

        for row in csv_reader:
            logger.info("Getting Agent ID for Agent UUID: %s", row[0])
            url = hostname.get() + f"/web/api/{api_version}/agents"
            agent_id = ""
            agent_config = ""
            param = {"uuid": row[0]}
            response = requests.get(
                url,
                params=param,
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            logger.debug(
                "Making API Call with the following:\nURL: %s\tParams: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                param,
                headers,
                proxy.get(),
                useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to get details for Agent UUID: %s Error code: %s Description: %s",
                    row[0],
                    str(response.status_code),
                    str(response.text),
                )
                continue
            else:
                r = response.json()
                agent_id = r["data"][0]["id"]
                logger.info(
                    "Successfully retrieved Agent ID: %s for Agent UUID: %s",
                    agent_id,
                    row[0],
                )

            logger.info("Getting Agent Config for Agent ID: %s", agent_id)
            url = (
                hostname.get()
                + f"/web/api/{api_version}/private/agents/{agent_id}/support-actions/configuration"
            )

            response = requests.get(
                url,
                params={},
                headers=headers,
                proxies={"http": proxy.get(), "https": proxy.get()},
                verify=useSSL.get(),
            )
            logger.debug(
                "Making API Call with the following:\nURL: %s\tParams: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                param,
                headers,
                proxy.get(),
                useSSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to get local config for Agent ID: %s Error code: %s Description: %s",
                    agent_id,
                    str(response.status_code),
                    str(response.text),
                )
                continue
            else:
                r = response.json()
                agent_config = r["data"]["configuration"]
                logger.info(
                    "Successfully retrieved local config for Agent ID: %s", agent_id
                )

            formatted_data = json.loads(agent_config)
            with open(json_file, "a+", encoding="utf-8") as f:
                logger.info("Writing local config for %s to %s", agent_id, json_file)
                f.write(f"\n{agent_id} - {row[0]}:\n")
                json.dump(formatted_data, f, indent=4)
                f.write("\n")

    logger.info("Done! Output file is - %s\n", json_file)


def exportUsers():
    """Function to handle getting User Details and writing to CSV or XLSX"""
    st = ScrolledText.ScrolledText(master=exportUsersFrame, state="disabled", height=10)
    st.configure(font=st_font)
    st.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    logger.debug("User selected %s file type", userOutputType.get())
    output_file_name = f"Export_Users_{datestamp}"
    csv_file = output_file_name + ".csv"
    xlsx_file = output_file_name + ".xlsx"

    COL_NAMES = [
        "Full Name",
        "Email",
        "Verified Email",
        "User ID",
        "Date Joined",
        "First Login Date",
        "Last Login",
        "2FA Enabled?",
        "2FA Method",
        "Lowest Role",
        "Scope",
        "Scope Roles",
        "Site Roles",
        "Tenant Roles",
        "API Token Dates",
        "Read-Only Groups",
        "Read-Only Email",
        "Read-Only Full Name",
        "Source",
        "Is System?",
    ]

    url = (
        hostname.get()
        + f"/web/api/{api_version}/users?{query_limits}&sortOrder=asc&sortBy=email"
    )
    first_run = True
    total_users = 0
    users = {}

    logger.info("Getting Users list")
    while url:
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": proxy.get(), "https": proxy.get()},
            verify=useSSL.get(),
        )
        logger.debug(
            "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            proxy.get(),
            useSSL.get(),
        )
        if response.status_code != 200:
            logger.error(
                "Failed to get users. Error code: %s Description: %s",
                str(response.status_code),
                str(response.text),
            )
            break
        else:
            data = response.json()
            cursor = data["pagination"]["nextCursor"]
            users = data["data"]
            if first_run:
                total_users = data["pagination"]["totalItems"]
                logger.debug(
                    "First run through user data. Total users in complete date set: %s",
                    total_users,
                )

            with open(csv_file, mode="a+", newline="", encoding="utf-8") as file:
                logger.debug("Writing CSV with User data")
                fieldnames = COL_NAMES
                csv_writer = csv.DictWriter(file, delimiter=",", fieldnames=fieldnames)
                if first_run:
                    csv_writer.writeheader()
                    logger.debug(
                        "First run through the data set complete, setting first run to False"
                    )
                    first_run = False

                for user in users:
                    logger.debug(
                        "Adding %s - %s to %s",
                        user["fullName"],
                        user["email"],
                        csv_file,
                    )
                    csv_writer.writerow(
                        {
                            "Full Name": user["fullName"],
                            "Email": user["email"],
                            "Verified Email": user["emailVerified"],
                            "User ID": user["id"],
                            "Date Joined": user["dateJoined"],
                            "First Login Date": user.get("firstLogin") or "Never",
                            "Last Login": user.get("lastLogin") or "Never",
                            "2FA Enabled?": user["twoFaEnabled"],
                            "2FA Method": user.get("primaryTwoFaMethod") or "N/A",
                            "Lowest Role": user["lowestRole"],
                            "Scope": user["scope"],
                            "Scope Roles": user["scopeRoles"],
                            "Site Roles": user.get("siteRoles") or "N/A",
                            "Tenant Roles": user.get("tenantRoles") or "N/A",
                            "API Token Dates": user.get("apiToken") or "N/A",
                            "Read-Only Groups": user["groupsReadOnly"],
                            "Read-Only Email": user["emailReadOnly"],
                            "Read-Only Full Name": user["fullNameReadOnly"],
                            "Source": user["source"],
                            "Is System?": user["isSystem"],
                        }
                    )

            if cursor:
                paramsnext = f"/web/api/{api_version}/users?{query_limits}&sortOrder=asc&sortBy=email&cursor={cursor}"
                url = hostname.get() + paramsnext
                logger.debug("Next cursor found, updating URL: %s", url)
            else:
                logger.debug("No cursor found, setting URL to None")
                url = None

    if userOutputType.get() == "xlsx":
        logger.debug("Creating new XLSX: %s", xlsx_file)
        workbook = Workbook(xlsx_file)
        logger.debug("Adding new worksheet: 'Users'")
        worksheet = workbook.add_worksheet("Users")
        if os.path.isfile(csv_file):
            with open(csv_file, "r", encoding="utf8") as f:
                logger.debug("Reading %s and writing to %s", csv_file, workbook)
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        worksheet.write(r, c, col)
            logger.debug("Deleting %s", csv_file)
            os.remove(csv_file)
        else:
            logger.error("%s not found.", csv_file)
        logger.debug("Closing XLSX")
        workbook.close()

    logger.info(
        "Done! Output file is - %s.%s\n", output_file_name, userOutputType.get()
    )


def export_ranger():
    """Function to handle exporting Ranger Inventory to CSV"""
    st = ScrolledText.ScrolledText(
        master=exportRangerInvFrame, state="disabled", height=10
    )
    st.configure(font=st_font)
    st.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(st)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    export_scope = exportRangerScope.get()
    ranger_time_period = exportRangerTimePeriod.get()
    if export_scope == "sites":
        scope_param = "siteIds"
    else:
        scope_param = "accountIds"
    if not inputcsv.get():
        logger.error("Must select a CSV containing Account or Site IDs")

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d")

    logger.debug(
        "Input options:\n\tScope: %s\n\tScope ID CSV: %s\n\tTime Period: %s",
        export_scope,
        inputcsv.get(),
        ranger_time_period,
    )

    with open(inputcsv.get()) as csv_file:
        logger.debug("Reading CSV: %s", inputcsv.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        for row in csv_reader:
            logger.info(
                "Exporting Ranger Inventory for %s scope ID: %s",
                export_scope.capitalize(),
                row[0],
            )
            firstrun = True
            endpoint = f"/web/api/{api_version}/ranger/table-view?{query_limits}&period={ranger_time_period}&{scope_param}={row[0]}"
            url = hostname.get() + endpoint
            while url:
                logger.debug(
                    "Making API Call with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy.get(),
                    useSSL.get(),
                )
                response = requests.get(
                    url,
                    headers=headers,
                    proxies={"http": proxy.get(), "https": proxy.get()},
                    verify=useSSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Status: %s Problem with the request. Details - %s ",
                        str(response.status_code),
                        str(response.text),
                    )
                    break
                else:
                    data = response.json()
                    cursor = data["pagination"]["nextCursor"]
                    data = data["data"]
                    if not data:
                        logger.info("No Ranger Inventory data returned. Exiting")
                        url = None
                        continue
                    else:
                        csv_filename = f"Ranger_Export-{export_scope.capitalize()}_{row[0]}_{ranger_time_period}_{datestamp}.csv"
                        logger.debug("Opening %s to write", csv_filename)
                        f = csv.writer(
                            open(csv_filename, "a+", newline="", encoding="utf-8")
                        )
                        if firstrun:
                            tmp = []
                            for key, value in data[0].items():
                                tmp.append(key)
                            logger.debug("Writing first row to %s", csv_filename)
                            f.writerow(tmp)
                            firstrun = False
                        for item in data:
                            tmp = []
                            for key, value in item.items():
                                tmp.append(value)
                            logger.debug("Writing data to %s", csv_filename)
                            f.writerow(tmp)
                    if cursor:
                        paramsnext = endpoint + f"&cursor={cursor}"
                        url = hostname.get() + paramsnext
                        logger.debug("Next cursor found, updating URL: %s", url)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None
                logger.info("Finished writing to %s", csv_filename)
        logger.info("Done exporting Ranger Inventory.")


def selectCSVFile():
    """Basic function to present user with browse window to source a CSV file for input"""
    file = tkinter.filedialog.askopenfilename()
    inputcsv.set(file)


# Login Menu Frame #############################
tk.Label(master=loginMenuFrame, image=logo).grid(row=0, column=0, columnspan=1, pady=20)

consoleAddressLabel = tk.Label(
    master=loginMenuFrame,
    text="Management Console URL:",
)
consoleAddressLabel.grid(row=1, column=0, pady=2)

consoleAddressEntry = ttk.Entry(master=loginMenuFrame, width=80)
consoleAddressEntry.grid(row=2, column=0, pady=2)

apikTokenLabel = tk.Label(master=loginMenuFrame, text="API Token:")
apikTokenLabel.grid(row=3, column=0, pady=(10, 2))

apikTokenEntry = ttk.Entry(master=loginMenuFrame, width=80)
apikTokenEntry.grid(row=4, column=0, pady=2)

tk.Label(
    master=loginMenuFrame,
    text="*API Token provided must have sufficient permissions to perform a given action.",
    font=frame_subnote_font,
).grid(row=5, column=0, pady=5)

proxyLabel = tk.Label(
    master=loginMenuFrame,
    text="Proxy (if required):",
)
proxyLabel.grid(row=6, column=0, pady=(10, 2))

proxyEntry = ttk.Entry(master=loginMenuFrame, width=80)

proxyEntry.grid(row=7, column=0, pady=2)

useSSLSwitch = ttk.Checkbutton(
    master=loginMenuFrame,
    text="Use SSL",
    style="Switch",
    variable=useSSL,
    onvalue=True,
    offvalue=False,
)
useSSLSwitch.grid(row=8, column=0, pady=10)

loginButton = ttk.Button(master=loginMenuFrame, text="Login", command=login)
loginButton.grid(row=9, column=0, columnspan=2, ipady=5, pady=10)

if LOG_LEVEL == logging.DEBUG:
    ttk.Label(
        master=loginMenuFrame,
        text=f"S1 Manager launched with --debug. Be sure to delete {LOG_NAME} when finished.",
        font=frame_subnote_font,
        foreground=frame_note_fg_color,
    ).grid(row=10, column=0, pady=10, ipadx=5, ipady=5)

tk.Label(
    master=loginMenuFrame,
    text=f"SentinelOne API: {api_version}\tS1 Manager: v{__version__}",
).grid(row=12, column=0, pady=(10, 5), sticky="s")
loginMenuFrame.pack()

# Main Menu Frame #############################
tk.Label(master=mainMenuFrame, image=logo).grid(row=0, column=0, columnspan=4, pady=20)

# Export - Column 0
ttk.Label(
    master=mainMenuFrame, text="Export Operations", font=frame_subtitle_font_underline
).grid(row=1, column=0, columnspan=2, pady=20)
ttk.Button(
    master=mainMenuFrame,
    text="Export Deep Visiblity Events",
    command=partial(switchFrames, exportFromDVFrame),
    width=32,
).grid(row=2, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Activity Log",
    command=partial(switchFrames, exportActivityLogFrame),
    width=32,
).grid(row=3, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Endpoints",
    command=partial(switchFrames, exportEndpointsFrame),
    width=32,
).grid(row=4, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Exclusions",
    command=partial(switchFrames, exportExclusionsFrame),
    width=32,
).grid(row=5, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Endpoint Tags",
    command=partial(switchFrames, exportEndpointTagsFrame),
    width=32,
).grid(row=6, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Local Config",
    command=partial(switchFrames, exportLocalConfigFrame),
    width=32,
).grid(row=2, column=1, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Users",
    command=partial(switchFrames, exportUsersFrame),
    width=32,
).grid(row=3, column=1, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Export Ranger Inventory",
    command=partial(switchFrames, exportRangerInvFrame),
    width=32,
).grid(row=4, column=1, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=5, column=1, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=6, column=1, sticky="ew", ipady=5, pady=5, padx=5)

# Manage - Column 2
tk.Label(
    master=mainMenuFrame, text="Manage Operations", font=frame_subtitle_font_underline
).grid(row=1, column=2, columnspan=2, pady=20)
ttk.Button(
    master=mainMenuFrame,
    text="Upgrade Agents",
    command=partial(switchFrames, upgradeFromCSVFrame),
    width=32,
).grid(row=2, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Move Agents",
    command=partial(switchFrames, moveAgentsFrame),
    width=32,
).grid(row=3, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Assign Customer Identifier",
    command=partial(switchFrames, assignCustomerIdentifierFrame),
    width=32,
).grid(row=4, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Decommission Agents",
    command=partial(switchFrames, decommissionAgentsFrame),
    width=32,
).grid(row=5, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=mainMenuFrame,
    text="Manage Endpoint Tags",
    command=partial(switchFrames, manageEndpointTagsFrame),
    width=32,
).grid(row=6, column=2, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=2, column=3, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=3, column=3, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=4, column=3, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=5, column=3, sticky="ew", ipady=5, pady=5, padx=5)
# ttk.Button(
#     master=mainMenuFrame,
#     text="-",
#     state=tk.DISABLED,
#     width=32,
# ).grid(row=6, column=3, sticky="ew", ipady=5, pady=5, padx=5)

if LOG_LEVEL == logging.DEBUG:
    ttk.Label(
        master=mainMenuFrame,
        text=f"S1 Manager launched with --debug. Be sure to delete {LOG_NAME} when finished.",
        font=frame_subnote_font,
        foreground=frame_note_fg_color,
    ).grid(row=10, column=0, columnspan=4, pady=10, ipadx=5, ipady=5)

tk.Label(
    master=mainMenuFrame,
    text="Note: Many of the processes can take a while to run. Be patient.",
    font=frame_subnote_font,
).grid(row=11, column=0, columnspan=4, padx=20, pady=20, sticky="s")


# Export from DV Frame #############################
tk.Label(
    master=exportFromDVFrame,
    text="Export Deep Visiblity Events",
    font=frame_title_font,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportFromDVFrame,
    text="Export Deep Visibility events to an XLSX by query ID as reference",
    font=frame_subtitle_font,
).grid(row=1, column=0, padx=20, pady=10)
tk.Label(master=exportFromDVFrame, text="1. Input Deep Visibility Query ID").grid(
    row=2, column=0, pady=2
)
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


# Search and Export Activity Log Frame #############################
tk.Label(
    master=exportActivityLogFrame,
    text="Search and Export Activity Log",
    font=frame_title_font,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportActivityLogFrame,
    text="Search Management Console Activity log and export results.",
    font=frame_subtitle_font,
).grid(row=1, column=0, padx=20, pady=10)
tk.Label(master=exportActivityLogFrame, text="1. Input FROM date (yyyy-mm-dd)").grid(
    row=2, column=0, pady=2
)
dateFrom = fromDateEntry = ttk.Entry(master=exportActivityLogFrame, width=40)
fromDateEntry.grid(row=3, column=0, pady=10)
tk.Label(master=exportActivityLogFrame, text="2. Input TO date (yyyy-mm-dd)").grid(
    row=4, column=0, pady=2
)
dateTo = toDateEntry = ttk.Entry(master=exportActivityLogFrame, width=40)
toDateEntry.grid(row=5, column=0, pady=10)
tk.Label(master=exportActivityLogFrame, text="3. Input search string").grid(
    row=6, column=0, pady=2
)
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
    font=frame_title_font,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=upgradeFromCSVFrame,
    text="Upgrade Agents to a specific package version by ID.",
    font=frame_subtitle_font,
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=upgradeFromCSVFrame, text="1. Export Packages List to source Package ID"
).grid(row=2, column=0, padx=20, pady=2)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Export Packages List",
    command=partial(upgradeFromCSV, True),
).grid(row=3, column=0, pady=10)
tk.Label(master=upgradeFromCSVFrame, text="2. Insert the Package ID").grid(
    row=4, column=0, pady=2
)
packageIDEntry = ttk.Entry(master=upgradeFromCSVFrame, width=80)
packageIDEntry.grid(row=5, column=0, pady=2)
tk.Label(
    master=upgradeFromCSVFrame,
    text="3. Select a CSV file containing a single column of endpoint names to upgrade",
).grid(row=6, column=0, padx=20, pady=2)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=7, column=0, pady=2)
tk.Label(master=upgradeFromCSVFrame, textvariable=inputcsv).grid(
    row=8, column=0, pady=2
)
useScheduleSwitch = ttk.Checkbutton(
    master=upgradeFromCSVFrame,
    text="Use Schedule",
    style="Switch",
    variable=useSchedule,
    onvalue=True,
    offvalue=False,
)
useScheduleSwitch.grid(row=9, column=0, pady=10)
tk.Label(
    master=upgradeFromCSVFrame,
    text="Note: Will request upgrade immediately, unless 'Use Schedule' is toggled on.",
    font=frame_subnote_font,
).grid(row=10, column=0, pady=2)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Submit",
    command=partial(upgradeFromCSV, False),
).grid(row=11, column=0, pady=10)
ttk.Button(
    master=upgradeFromCSVFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=12, column=0, ipadx=10, pady=10)


# Move Agents Frame #############################
tk.Label(
    master=moveAgentsFrame,
    text="Move Agents",
    font=frame_title_font,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=moveAgentsFrame,
    text="Move Agents to specified Site ID and Group ID.",
    font=frame_subtitle_font,
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=moveAgentsFrame,
    text="If the target group is dynamic, the agent will be moved to the site only.",
).grid(row=2, column=0, pady=2)
tk.Label(master=moveAgentsFrame, text="1. Export Groups List to get group IDs").grid(
    row=3, column=0, pady=2
)
ttk.Button(
    master=moveAgentsFrame,
    text="Export Groups List",
    command=partial(moveAgents, True),
).grid(row=4, column=0, pady=10)
tk.Label(
    master=moveAgentsFrame,
    text="2. Select a CSV file constructed of three columns:\nendpoints names, target group IDs, target site IDs",
).grid(row=5, column=0, padx=20, pady=10)
ttk.Button(master=moveAgentsFrame, text="Browse", command=selectCSVFile).grid(
    row=6, column=0, pady=10
)
tk.Label(master=moveAgentsFrame, textvariable=inputcsv).grid(row=7, column=0, pady=10)
ttk.Button(
    master=moveAgentsFrame,
    text="Submit",
    command=partial(moveAgents, False),
).grid(row=8, column=0, pady=10)
ttk.Button(
    master=moveAgentsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=9, column=0, ipadx=10, pady=10)


# Assign Customer Identifier Frame #############################
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="Assign Customer Identifier",
    font=frame_title_font,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="Assign a Customer Identifier to one or more Agents.",
    font=frame_subtitle_font,
).grid(row=1, column=0, pady=2)
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="1. Input the Customer Identifier to assign",
).grid(row=2, column=0, padx=20, pady=2)
customerIdentifierEntry = ttk.Entry(master=assignCustomerIdentifierFrame, width=80)
customerIdentifierEntry.grid(row=3, column=0, pady=(2, 10))
tk.Label(
    master=assignCustomerIdentifierFrame,
    text="2. Select a CSV file containing a single column with endpoint names",
).grid(row=4, column=0, padx=20, pady=2)
ttk.Button(
    master=assignCustomerIdentifierFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=5, column=0, pady=10)
tk.Label(master=assignCustomerIdentifierFrame, textvariable=inputcsv).grid(
    row=6, column=0, pady=10
)
ttk.Button(
    master=assignCustomerIdentifierFrame,
    text="Submit",
    command=assignCustomerIdentifier,
).grid(row=7, column=0, pady=10)
ttk.Button(
    master=assignCustomerIdentifierFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=8, column=0, ipadx=10, pady=10)


# Decommission Agents from CSV Frame #############################
tk.Label(
    master=decommissionAgentsFrame,
    text="Decommission Agents",
    font=frame_title_font,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=decommissionAgentsFrame,
    text="1. Select a CSV file containing a single column of endpoint names to be decommissioned",
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=decommissionAgentsFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=2, column=0, pady=10)
tk.Label(master=decommissionAgentsFrame, textvariable=inputcsv).grid(
    row=3, column=0, pady=10
)
ttk.Button(
    master=decommissionAgentsFrame,
    text="Submit",
    command=decommissionAgents,
).grid(row=4, column=0, pady=10)
ttk.Button(
    master=decommissionAgentsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=5, column=0, ipadx=10, pady=10)


# Export all agents Frame #############################
tk.Label(
    master=exportEndpointsFrame,
    text="Export All Endpoints",
    font=frame_title_font,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=exportEndpointsFrame,
    text="Exports all Agent details to a CSV or XLSX",
    font=frame_subtitle_font,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
endpointOutputType = tk.StringVar()
endpointOutputType.set("csv")
ttk.Radiobutton(
    exportEndpointsFrame, text="CSV", variable=endpointOutputType, value="csv"
).grid(row=2, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    exportEndpointsFrame, text="XLSX", variable=endpointOutputType, value="xlsx"
).grid(row=2, column=1, padx=10, pady=2, sticky="w")
ttk.Button(
    master=exportEndpointsFrame,
    text="Export",
    command=exportAllAgents,
).grid(row=3, column=0, columnspan=2, pady=10)
ttk.Button(
    master=exportEndpointsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=4, column=0, columnspan=2, ipadx=10, pady=10)


# Export Exclusions #############################
tk.Label(
    master=exportExclusionsFrame, text="Export Exclusions", font=frame_title_font
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportExclusionsFrame,
    text="Exports all Exclusions to an XLSX",
    font=frame_subtitle_font,
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


# Export Endpoint Tag IDs Frame #############################
tk.Label(
    master=exportEndpointTagsFrame, text="Export Endpoint Tags", font=frame_title_font
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportEndpointTagsFrame,
    text="Exports Endpoint Tag details to CSV for all scopes.",
    font=frame_subtitle_font,
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=exportEndpointTagsFrame,
    text="Export",
    command=exportEndpointTags,
).grid(row=2, column=0, pady=10)
ttk.Button(
    master=exportEndpointTagsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=3, column=0, ipadx=10, pady=10)


# Manage Endpoint Tags Frame #############################
tk.Label(
    master=manageEndpointTagsFrame, text="Manage Endpoint Tags", font=frame_title_font
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=manageEndpointTagsFrame,
    text="Add or Remove Endpoint Tags from Agents.",
    font=frame_subtitle_font,
).grid(row=1, column=0, columnspan=2, pady=2)
tk.Label(master=manageEndpointTagsFrame, text="1. Select Action").grid(
    row=2, column=0, columnspan=2, padx=20, pady=2
)
endpointTagsAction = tk.StringVar()
endpointTagsAction.set("add")
ttk.Radiobutton(
    manageEndpointTagsFrame,
    text="Add Endpoint Tag",
    variable=endpointTagsAction,
    value="add",
).grid(row=3, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    manageEndpointTagsFrame,
    text="Remove Endpoint Tag",
    variable=endpointTagsAction,
    value="remove",
).grid(row=3, column=1, padx=10, pady=2, sticky="w")
tk.Label(master=manageEndpointTagsFrame, text="2. Input Endpoint Tag ID").grid(
    row=4, column=0, columnspan=2, padx=20, pady=2
)
tagIDEntry = ttk.Entry(master=manageEndpointTagsFrame, width=80)
tagIDEntry.grid(row=5, column=0, columnspan=2, pady=(2, 10))
tk.Label(
    master=manageEndpointTagsFrame,
    text="3. Select Agent Identifier type. This should align with your source CSV.",
).grid(row=6, column=0, columnspan=2, padx=20, pady=2)
agentIDType = tk.StringVar()
agentIDType.set("uuid")
ttk.Radiobutton(
    manageEndpointTagsFrame, text="Agent UUID", variable=agentIDType, value="uuid"
).grid(row=7, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    manageEndpointTagsFrame, text="Endpoint Name", variable=agentIDType, value="name"
).grid(row=7, column=1, padx=10, pady=2, sticky="w")
tk.Label(
    master=manageEndpointTagsFrame,
    text="4. Select a CSV file containing a single column of values (uuids or endpoint names)",
).grid(row=8, column=0, columnspan=2, padx=20, pady=2)
ttk.Button(
    master=manageEndpointTagsFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=9, column=0, columnspan=2, pady=10)
tk.Label(master=manageEndpointTagsFrame, textvariable=inputcsv).grid(
    row=10, column=0, columnspan=2, pady=10
)
ttk.Button(
    master=manageEndpointTagsFrame, text="Submit", command=manageEndpointTags
).grid(row=11, column=0, columnspan=2, pady=10)
ttk.Button(
    master=manageEndpointTagsFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=12, column=0, columnspan=2, ipadx=10, pady=10)


# Export Agent Local Config Frame #############################
tk.Label(
    master=exportLocalConfigFrame, text="Export Local Config", font=frame_title_font
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=exportLocalConfigFrame,
    text="Exports the local agent configuration to a single JSON file.",
    font=frame_subtitle_font,
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=exportLocalConfigFrame,
    text="1. Select a CSV file containing a single column of agent UUIDs",
).grid(row=2, column=0, padx=20, pady=2)
ttk.Button(
    master=exportLocalConfigFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=3, column=0, pady=10)
tk.Label(master=exportLocalConfigFrame, textvariable=inputcsv).grid(
    row=4, column=0, pady=10
)
ttk.Button(
    master=exportLocalConfigFrame,
    text="Export",
    command=exportLocalConfig,
).grid(row=5, column=0, pady=10)
ttk.Button(
    master=exportLocalConfigFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=6, column=0, ipadx=10, pady=10)


# Export Users Frame #############################
tk.Label(master=exportUsersFrame, text="Export Users", font=frame_title_font).grid(
    row=0, column=0, columnspan=2, padx=20, pady=20
)
tk.Label(
    master=exportUsersFrame,
    text="Exports User details to CSV.",
    font=frame_subtitle_font,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
userOutputType = tk.StringVar()
userOutputType.set("csv")
ttk.Radiobutton(
    exportUsersFrame, text="CSV", variable=userOutputType, value="csv"
).grid(row=2, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    exportUsersFrame, text="XLSX", variable=userOutputType, value="xlsx"
).grid(row=2, column=1, padx=10, pady=2, sticky="w")
ttk.Button(
    master=exportUsersFrame,
    text="Export",
    command=exportUsers,
).grid(row=3, column=0, columnspan=2, pady=10)
ttk.Button(
    master=exportUsersFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=4, column=0, columnspan=2, ipadx=10, pady=10)


# Export Ranger Inventory Frame #############################
tk.Label(
    master=exportRangerInvFrame, text="Export Ranger Inventory", font=frame_title_font
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=exportRangerInvFrame,
    text="Exports Ranger Inventory details to CSV",
    font=frame_subtitle_font,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
tk.Label(
    master=exportRangerInvFrame,
    text="1. Select which scope type to export Ranger Inventory from.",
).grid(row=2, column=0, columnspan=2, padx=20, pady=2)
exportRangerScope = tk.StringVar()
exportRangerScope.set("accounts")
ttk.Radiobutton(
    exportRangerInvFrame,
    text="Account",
    variable=exportRangerScope,
    value="accounts",
).grid(row=3, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    exportRangerInvFrame, text="Site", variable=exportRangerScope, value="sites"
).grid(row=3, column=1, padx=10, pady=2, sticky="w")
tk.Label(
    master=exportRangerInvFrame,
    text="2. Select a CSV containing a single column of Account or Site IDs.",
).grid(row=4, column=0, columnspan=2, padx=20, pady=2)
ttk.Button(
    master=exportRangerInvFrame,
    text="Browse",
    command=selectCSVFile,
).grid(row=5, column=0, columnspan=2, pady=2)
tk.Label(master=exportRangerInvFrame, textvariable=inputcsv).grid(
    row=6, column=0, columnspan=2, pady=2
)
tk.Label(
    master=exportRangerInvFrame,
    text="3. Specify time period for data export",
).grid(row=7, column=0, columnspan=2, padx=20, pady=2)
available_timeperiods = ("", "latest", "last12h", "last24h", "last3d", "last7d")
exportRangerTimePeriod = tk.StringVar()
exportRangerTimePeriod.set(available_timeperiods[1])
ttk.OptionMenu(
    exportRangerInvFrame, exportRangerTimePeriod, *available_timeperiods
).grid(row=8, column=0, columnspan=2, pady=10)
ttk.Button(
    master=exportRangerInvFrame,
    text="Export",
    command=export_ranger,
).grid(row=9, column=0, columnspan=2, pady=10)
ttk.Button(
    master=exportRangerInvFrame,
    text="Back to Main Menu",
    command=goBacktoMainPage,
).grid(row=10, column=0, columnspan=2, ipadx=10, pady=10)


window.mainloop()
