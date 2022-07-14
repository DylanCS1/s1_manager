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
from pathlib import Path
from tkinter import UNDERLINE, ttk

import aiohttp
import requests
from PIL import Image, ImageTk
from xlsxwriter.workbook import Workbook

# CONSTS
__version__ = "2022.2.0"
API_VERSION = "v2.1"
DIR_PATH = os.path.dirname(os.path.realpath(__file__))
QUERY_LIMITS = "limit=1000"
headers = {}

# LOG SETTINGS
if len(sys.argv) > 1 and sys.argv[1] == "--debug":
    LOG_LEVEL = logging.DEBUG
    LOG_NAME = f"s1_manager_debug_{datetime.datetime.now().strftime('%Y-%m-%d')}_{__version__}.log"
else:
    LOG_LEVEL = logging.INFO
    LOG_NAME = (
        f"s1_manager_{datetime.datetime.now().strftime('%Y-%m-%d')}_{__version__}.log"
    )
if LOG_LEVEL == logging.DEBUG:
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)s - %(message)s"
else:
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

# WINDOW SETTINGS
window = tk.Tk()
window.title("S1 Manager")
if platform.system() == "Windows":
    window.iconbitmap(os.path.join(DIR_PATH, "ico/s1_manager.ico"))
window.minsize(900, 700)

# THEME
window.tk.call("source", os.path.join(DIR_PATH, "theme/forest-dark.tcl"))
LOGO = ImageTk.PhotoImage(Image.open(os.path.join(DIR_PATH, "ico/s1_manager.png")))
ttk.Style().theme_use("forest-dark")
FRAME_TITLE_FONT = ("Courier", 24, UNDERLINE)
FRAME_SUBTITLE_FONT_UNDERLINE = ("Arial", 14, UNDERLINE)
FRAME_SUBTITLE_FONT = ("Arial", 12)
FRAME_SUBNOTE_FONT = ("Arial", 10)
FRAME_NOTE_FG_COLOR = "red"
ST_FONT = "TkFixedFont"

# FRAME CONSTS
LOGIN_MENU_FRAME = ttk.Frame()
MAIN_MENU_FRAME = ttk.Frame()
EXPORT_FROM_DV_FRAME = ttk.Frame()
EXPORT_ACTIVITY_LOG_FRAME = ttk.Frame()
EXPORT_ENDPOINTS_FRAME = ttk.Frame()
EXPORT_ENDPOINT_TAGS_FRAME = ttk.Frame()
EXPORT_EXCLUSIONS_FRAME = ttk.Frame()
EXPORT_LOCAL_CONFIG_FRAME = ttk.Frame()
EXPORT_USERS_FRAME = ttk.Frame()
EXPORT_RANGER_INV_FRAME = ttk.Frame()
UPGRADE_FROM_CSV_FRAME = ttk.Frame()
MOVE_AGENTS_FRAME = ttk.Frame()
ASSIGN_CUSTOMER_ID_FRAME = ttk.Frame()
DECOMMISSION_AGENTS_FRAME = ttk.Frame()
MANAGE_ENDPOINT_TAGS_FRAME = ttk.Frame()
BULK_RESOLVE_THREATS_FRAME = ttk.Frame()
UPDATE_SYSTEM_CONFIG_FRAME = ttk.Frame()
BULK_ENABLE_AGENTS_FRAME = ttk.Frame()
ERROR = tk.StringVar()
HOSTNAME = tk.StringVar()
API_TOKEN = tk.StringVar()
PROXY = tk.StringVar()
INPUT_FILE = tk.StringVar()
USE_SSL = tk.BooleanVar()
USE_SSL.set(True)
USE_SCHEDULE = tk.BooleanVar()
USE_SCHEDULE.set(False)


class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget
    Adapted from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06"""

    def __init__(self, text):
        logging.Handler.__init__(self)
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state="normal")
            self.text.insert(tk.END, msg + "\n")
            self.text.configure(state="disabled")
            self.text.yview(tk.END)

        self.text.after(0, append)


# Helper Functions
def test_login(hostname, apitoken, proxy):
    """Function to test login using APIToken or Token"""

    headers = {
        "Content-type": "application/json",
        "Authorization": "ApiToken " + apitoken,
        "Accept": "application/json",
    }
    response = requests.get(
        hostname + f"/web/api/{API_VERSION}/system/info",
        headers=headers,
        proxies={"http": proxy, "https": proxy},
        verify=USE_SSL.get(),
    )

    if response.status_code == 200:
        return headers, True
    else:
        headers = {
            "Content-type": "application/json",
            "Authorization": "Token " + apitoken,
        }
        response = requests.get(
            hostname + f"/web/api/{API_VERSION}/system/info",
            headers=headers,
            proxies={"http": proxy, "https": proxy},
            verify=USE_SSL.get(),
        )
        response.raise_for_status()
        if response.status_code == 200:
            return headers, True
        else:
            return 0, False


def login():
    """Function to handle login actions"""
    HOSTNAME.set(console_address_entry.get())
    API_TOKEN.set(api_token_entry.get())
    PROXY.set(proxy_entry.get())
    global headers

    if not HOSTNAME.get() or not API_TOKEN.get():
        tk.Label(
            master=LOGIN_MENU_FRAME,
            text="'Management Console URL' and 'API Token' cannot be empty.",
            fg=FRAME_NOTE_FG_COLOR,
            font=FRAME_SUBNOTE_FONT,
        ).grid(row=11, column=0, columnspan=2, pady=10)
    else:
        headers, login_succ = test_login(HOSTNAME.get(), API_TOKEN.get(), PROXY.get())
        if login_succ:
            LOGIN_MENU_FRAME.pack_forget()
            MAIN_MENU_FRAME.pack()
        else:
            tk.Label(
                master=LOGIN_MENU_FRAME,
                text="Authentication failed. Please check credentials and try again",
                fg=FRAME_NOTE_FG_COLOR,
                font=FRAME_SUBNOTE_FONT,
            ).grid(row=11, column=0, columnspan=2, pady=10)


def go_back_to_mainpage():
    """Function to handle moving back to the Main Menu Frame"""
    _list = window.winfo_children()
    for item in _list:
        if item.winfo_children():
            _list.extend(item.winfo_children())
    for item in _list:
        if isinstance(item, tk.Toplevel) is not True:
            item.pack_forget()
    MAIN_MENU_FRAME.pack()


def switch_frames(framename):
    """Function to handle switching tkinter frames"""
    INPUT_FILE.set("")
    MAIN_MENU_FRAME.pack_forget()
    framename.pack()


def select_csv_file():
    """Basic function to present user with browse window to source a CSV file for input"""
    file = tkinter.filedialog.askopenfilename()
    INPUT_FILE.set(file)


# Tool operation functions
def export_from_dv():
    """Function to export events from Deep Visibility by DV query ID"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_FROM_DV_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
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
        params = f"/web/api/{API_VERSION}/dv/events/{querytype}?queryId={dv_query_id}"
        url = hostname + params
        while url:
            async with session.get(
                url, headers=headers, proxy=proxy, ssl=USE_SSL.get()
            ) as response:
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    proxy,
                    USE_SSL.get(),
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
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "ip":
                                csv_file = csv.writer(
                                    open(dv_ip, "a+", newline="", encoding="utf-8")
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "url":
                                csv_file = csv.writer(
                                    open(dv_url, "a+", newline="", encoding="utf-8")
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "dns":
                                csv_file = csv.writer(
                                    open(dv_dns, "a+", newline="", encoding="utf-8")
                                )
                                if firstrun:
                                    tmp = []
                                    for key, value in data.items():
                                        tmp.append(key)
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "process":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "registry":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "scheduled_task":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrun = False
                                tmp = []
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)
                    if cursor:
                        paramsnext = f"/web/api/{API_VERSION}/dv/events/{querytype}?cursor={cursor}&queryId={dv_query_id}&{QUERY_LIMITS}"
                        url = hostname + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def run(hostname, dv_query_id, proxy):
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

    dv_query_id = query_id_entry.get()
    if dv_query_id:
        logger.info("Processing DV Query ID: %s", dv_query_id)
        dv_query_id = dv_query_id.split(",")
        if platform.system() == "Windows":
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        asyncio.run(run(HOSTNAME.get(), dv_query_id, PROXY.get()))
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
            worksheet = workbook.add_worksheet(csvfile.split(".", maxsplit=-1)[0])
            if os.path.isfile(csvfile):
                with open(csvfile, "r", encoding="utf8") as file:
                    logger.debug("Reading %s and writing to %s", csvfile, workbook)
                    reader = csv.reader(file)
                    for r_idx, row in enumerate(reader):
                        for c_idx, col in enumerate(row):
                            worksheet.write(r_idx, c_idx, col)
                logger.debug("Deleting %s", csvfile)
                os.remove(csvfile)
        workbook.close()
        logger.info("Done! Created the file %s\n", xlsx_filename)
    else:
        logger.error("Please enter a valid DV Query ID and try again.")


def export_activity_log(search_only):
    """Function to search for Activity events by date range or export Activity events"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_ACTIVITY_LOG_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    os.environ["TZ"] = "UTC"
    date_format = "%Y-%m-%d"
    fromdate_epoch = (
        str(int(time.mktime(time.strptime(date_from.get(), date_format)))) + "000"
    )
    todate_epoch = (
        str(int(time.mktime(time.strptime(date_to.get(), date_format)))) + "000"
    )
    logger.debug(
        "Input FROM Date: %s Input TO Date: %s", date_from.get(), date_to.get()
    )
    logger.debug(
        "Epoch-converted FROM Date: %s Epoch-converted TO Date: %s",
        fromdate_epoch,
        todate_epoch,
    )
    if date_from.get() and date_to.get():
        url = (
            HOSTNAME.get()
            + f"/web/api/{API_VERSION}/activities?{QUERY_LIMITS}&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&includeHidden=false"
        )
        logger.debug("Search only state: %s", search_only)
        if search_only:
            logger.info("Starting search for '%s'", string_search_entry.get())
            while url:
                response = requests.get(
                    url,
                    headers=headers,
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
                )
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
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
                                string_search_entry.get().upper()
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
                                    string_search_entry.get().upper()
                                    in item["secondaryDescription"].upper()
                                ):
                                    logger.info(
                                        "%s - %s - %s",
                                        item["createdAt"],
                                        item["primaryDescription"],
                                        item["secondaryDescription"],
                                    )
                    if cursor:
                        paramsnext = f"/web/api/{API_VERSION}/activities?{QUERY_LIMITS}&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&cursor={cursor}&includeHidden=false"
                        url = HOSTNAME.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None
        else:
            datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
            csv_filename = f"Activity_Log_Export_{datestamp}.csv"
            logger.info("Creating and opening %s", csv_filename)
            csv_file = csv.writer(
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
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
                )
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
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
                            csv_file.writerow(tmp)
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
                            csv_file.writerow(tmp)
                    if cursor:
                        paramsnext = f"/web/api/{API_VERSION}/activities?{QUERY_LIMITS}&createdAt__between={fromdate_epoch}-{todate_epoch}&countOnly=false&cursor={cursor}&includeHidden=false"
                        url = HOSTNAME.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None
            logger.info("Done! Output file is - %s\n", csv_filename)
    else:
        logger.error("You must state a FROM date and a TO date")


def upgrade_from_csv(just_packages):
    """Function to upgrade Agents via API"""
    scroll_text = ScrolledText.ScrolledText(
        master=UPGRADE_FROM_CSV_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    csv_filename = f"Available_Packages_List_{datestamp}.csv"

    logger.debug("Just packages set to: %s", just_packages)
    if just_packages:
        params = f"/web/api/{API_VERSION}/update/agent/packages?sortBy=updatedAt&sortOrder=desc&countOnly=false&{QUERY_LIMITS}"
        url = HOSTNAME.get() + params
        csv_file = csv.writer(open(csv_filename, "a+", newline="", encoding="utf-8"))
        csv_file.writerow(
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
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
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
                        csv_file.writerow(
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
                    paramsnext = f"/web/api/{API_VERSION}/update/agent/packages?sortBy=updatedAt&sortOrder=desc&{QUERY_LIMITS}&cursor={cursor}&countOnly=false"
                    url = HOSTNAME.get() + paramsnext
                    logger.debug("Found next cursor: %s", cursor)
                else:
                    logger.debug("No cursor found, setting URL to None")
                    url = None
        logger.info("SentinelOne agent packages list written to: %s", csv_filename)
    else:
        with open(INPUT_FILE.get(), encoding="utf-8") as csv_file:
            logger.debug("Reading CSV: %s", INPUT_FILE.get())
            csv_reader = csv.reader(csv_file, delimiter=",")
            line_count = 0
            logger.debug("Use Schedule value: %s", USE_SCHEDULE.get())
            for row in csv_reader:
                logger.info("Upgrading endpoint named - %s", row[0])
                url = (
                    HOSTNAME.get()
                    + f"/web/api/{API_VERSION}/agents/actions/update-software"
                )
                body = {
                    "filter": {"computerName": row[0]},
                    "data": {
                        "packageId": package_id_entry.get(),
                        "isScheduled": USE_SCHEDULE.get(),
                    },
                }
                response = requests.post(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
                )
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tData: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    json.dumps(body),
                    PROXY.get(),
                    USE_SSL.get(),
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
                    if USE_SCHEDULE.get():
                        logger.info(
                            "Upgrade should follow schedule defined in Management Console."
                        )
                line_count += 1
            if line_count < 1:
                logger.info("Finished! Input file %s was empty.", INPUT_FILE.get())
            else:
                logger.info("Finished! Processed %d lines.", line_count)


def move_agents(just_groups):
    """Function to move Agents using API"""
    scroll_text = ScrolledText.ScrolledText(
        master=MOVE_AGENTS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    if just_groups:
        logger.debug("Just move agents to groups: %s", just_groups)
        params = f"/web/api/{API_VERSION}/groups?isDefault=false&limit=200&type=static&countOnly=false"
        url = HOSTNAME.get() + params
        csv_filename = "Group_To_ID_Map.csv"
        csv_file = csv.writer(open(csv_filename, "a+", newline="", encoding="utf-8"))
        csv_file.writerow(["Name", "ID", "Site ID", "Created By"])
        while url:
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
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
                        csv_file.writerow(
                            [
                                [data["name"]],
                                data["id"],
                                data["siteId"],
                                data["creator"],
                            ]
                        )
                if cursor:
                    paramsnext = f"/web/api/{API_VERSION}/groups?isDefault=false&limit=200&type=static&cursor={cursor}&countOnly=false"
                    url = HOSTNAME.get() + paramsnext
                    logger.debug("Found next cursor: %s", cursor)
                else:
                    logger.debug("No cursor found, setting URL to None")
                    url = None
        logger.info("Added group mapping to the file %s", csv_filename)
    else:
        with open(INPUT_FILE.get(), encoding="utf-8") as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            line_count = 0
            for row in csv_reader:
                logger.info("Moving endpoint name %s to Site ID %s", row[0], row[2])
                url = (
                    HOSTNAME.get()
                    + f"/web/api/{API_VERSION}/agents/actions/move-to-site"
                )
                body = {
                    "filter": {"computerName": row[0]},
                    "data": {"targetSiteId": row[2]},
                }
                response = requests.post(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
                )
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tData: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    json.dumps(body),
                    PROXY.get(),
                    USE_SSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Failed to transfer endpoint %s to site %s Error code: %s Description: %s",
                        row[0],
                        row[2],
                        str(response.status_code),
                        str(response.text),
                    )
                    continue
                else:
                    data = response.json()
                    logger.info("Moved %s endpoints", data["data"]["affected"])
                logger.info("Moving endpoint name %s to Group ID %s", row[0], row[1])
                url = (
                    HOSTNAME.get()
                    + f"/web/api/{API_VERSION}/groups/"
                    + row[1]
                    + "/move-agents"
                )
                body = {"filter": {"computerName": row[0]}}
                response = requests.put(
                    url,
                    data=json.dumps(body),
                    headers=headers,
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Failed to transfer endpoint %s to group %s Error code: %s Description: %s",
                        row[0],
                        row[1],
                        str(response.status_code),
                        str(response.text),
                    )
                    continue
                else:
                    data = response.json()
                    logger.info("Moved %s endpoints", data["data"]["agentsMoved"])
                line_count += 1
            if line_count < 1:
                logger.info("Finished! Input file %s was empty.", INPUT_FILE.get())
            else:
                logger.info("Finished! Processed %d lines.", line_count)


def assign_customer_id():
    """Function to add a Customer Identifier to one or more Agents via API"""
    scroll_text = ScrolledText.ScrolledText(
        master=ASSIGN_CUSTOMER_ID_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    with open(INPUT_FILE.get(), encoding="utf-8") as csv_file:
        logger.debug("Reading CSV: %s", INPUT_FILE.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0
        for row in csv_reader:
            logger.info("Updating customer identifier for endpoint - %s", row[0])
            url = (
                HOSTNAME.get()
                + f"/web/api/{API_VERSION}/agents/actions/set-external-id"
            )
            body = {
                "filter": {"computerName": row[0]},
                "data": {"externalId": customer_id_entry.get()},
            }
            response = requests.post(
                url,
                data=json.dumps(body),
                headers=headers,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to update customer identifier for endpoint %s Error code: %s Description: %s",
                    row[0],
                    str(response.status_code),
                    str(response.text),
                )
            else:
                json_response = response.json()
                affected_num_of_endpoints = json_response["data"]["affected"]
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
            logger.info("Finished! Input file %s was empty.", INPUT_FILE.get())
        else:
            logger.info("Finished! Processed %d lines.", line_count)


def export_all_agents():
    """Function to export a list of all Agents and details to CSV"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_ENDPOINTS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    output_file_name = f"Export_Endpoints_{datestamp}"
    csv_filename = output_file_name + ".csv"
    xlsx_file = output_file_name + ".xlsx"

    url = HOSTNAME.get() + f"/web/api/{API_VERSION}/export/agents-light"

    logger.info("Starting to request endpoint data.")

    session = requests.Session()
    with session.get(
        url,
        headers=headers,
        proxies={"http": PROXY.get(), "https": PROXY.get()},
        verify=USE_SSL.get(),
        stream=True,
    ) as download:
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
        )
        download.raise_for_status()

        logger.info("Writing to %s", csv_filename)
        with open(csv_filename, mode="wb") as new_file:
            for chunk in download.iter_content(chunk_size=1024 * 1024):
                logger.debug(
                    "Writing chunk: %s", chunk
                )  # Super-noisy, only uncomment if absolutely necessary for troubleshooting.
                new_file.write(chunk)

    logger.info("Creating new XLSX: %s", xlsx_file)
    workbook = Workbook(xlsx_file)
    logger.debug("Adding new worksheet: 'Endpoints'")
    worksheet = workbook.add_worksheet("Endpoints")
    if os.path.isfile(csv_filename):
        with open(csv_filename, "r", encoding="utf8") as csv_file:
            logger.info("Reading %s and writing to %s", csv_filename, xlsx_file)
            reader = csv.reader(csv_file)
            for r_idx, row in enumerate(reader):
                for c_idx, col in enumerate(row):
                    worksheet.write(r_idx, c_idx, col)
        logger.debug("Deleting %s", csv_filename)
        os.remove(csv_filename)
    else:
        logger.error("%s not found.", csv_filename)
    logger.debug("Closing XLSX")
    workbook.close()

    logger.info("Done! Output file is - %s.%s\n", output_file_name)


def decommission_agents():
    """Function to decommission specified agents via API"""
    scroll_text = ScrolledText.ScrolledText(
        master=DECOMMISSION_AGENTS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    with open(INPUT_FILE.get(), encoding="utf-8") as csv_file:
        logger.debug("Reading CSV: %s", INPUT_FILE.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0
        for row in csv_reader:
            logger.info("Decommissioning Endpoint - %s", row[0])
            logger.info("Getting endpoint ID for %s", row[0])
            url = (
                HOSTNAME.get()
                + f"/web/api/{API_VERSION}/agents?countOnly=false&computerName={row[0]}&{QUERY_LIMITS}"
            )
            response = requests.get(
                url,
                headers=headers,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Failed to get ID for endpoint %s Error code: %s Description: %s",
                    row[0],
                    str(response.status_code),
                    str(response.text),
                )
            else:
                json_response = response.json()
                totalitems = json_response["pagination"]["totalItems"]
                logger.debug("Total items returned: %s", totalitems)
                if totalitems < 1:
                    logger.info(
                        "Could not locate any IDs for endpoint named %s - Please note the query is CaSe SenSiTiVe",
                        row[0],
                    )
                else:
                    json_response = json_response["data"]
                    uuidslist = []
                    for item in json_response:
                        uuidslist.append(item["id"])
                        logger.info(
                            "Found ID %s! Adding it to be decommissioned", item["id"]
                        )
                    url = (
                        HOSTNAME.get()
                        + f"/web/api/{API_VERSION}/agents/actions/decommission"
                    )
                    body = {"filter": {"ids": uuidslist}}
                    response = requests.post(
                        url,
                        data=json.dumps(body),
                        headers=headers,
                        proxies={"http": PROXY.get(), "https": PROXY.get()},
                        verify=USE_SSL.get(),
                    )
                    logger.debug(
                        "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                        url,
                        headers,
                        PROXY.get(),
                        USE_SSL.get(),
                    )
                    if response.status_code != 200:
                        logger.error(
                            "Failed to decommission endpoint %s Error code: %s Description: %s",
                            row[0],
                            str(response.status_code),
                            str(response.text),
                        )
                    else:
                        json_response = response.json()
                        affected_num_of_endpoints = json_response["data"]["affected"]
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
            logger.info("Finished! Input file %s was empty.", INPUT_FILE.get())
        else:
            logger.info("Finished! Processed %d lines.", line_count)


def export_exclusions():
    """Function to export Exclusions to CSV"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_EXCLUSIONS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    async def get_accounts(session):
        logger.info("Getting accounts data")
        params = (
            f"/web/api/{API_VERSION}/accounts?{QUERY_LIMITS}"
            + "&countOnly=false&tenant=true"
        )
        url = HOSTNAME.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=PROXY.get()) as response:
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
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
                        paramsnext = f"/web/api/{API_VERSION}/accounts?{QUERY_LIMITS}&cursor={cursor}&countOnly=false&tenant=true"
                        url = HOSTNAME.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def get_sites(session):
        logger.info("Getting sites data")
        params = (
            f"/web/api/{API_VERSION}/sites?{QUERY_LIMITS}&countOnly=false&tenant=true"
        )
        url = HOSTNAME.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=PROXY.get()) as response:
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
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
                        paramsnext = f"/web/api/{API_VERSION}/sites?{QUERY_LIMITS}&cursor={cursor}&countOnly=false&tenant=true"
                        url = HOSTNAME.get() + paramsnext
                        logger.debug("Found next cursor: %s", cursor)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None

    async def get_groups(session):
        logger.info("Getting groups data")
        params = (
            f"/web/api/{API_VERSION}/groups?{QUERY_LIMITS}&countOnly=false&tenant=true"
        )
        url = HOSTNAME.get() + params
        while url:
            async with session.get(url, headers=headers, proxy=PROXY.get()) as response:
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
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
                        paramsnext = f"/web/api/{API_VERSION}/groups?{QUERY_LIMITS}&cursor={cursor}&countOnly=false&tenant=true"
                        url = HOSTNAME.get() + paramsnext
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
        params = f"/web/api/{API_VERSION}/exclusions?{QUERY_LIMITS}&type={querytype}&countOnly=false"
        url = HOSTNAME.get() + params + exparam
        while url:
            async with session.get(url, headers=headers, proxy=PROXY.get()) as response:
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
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
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrunpath = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "certificate":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstruncert = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "browser":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrunbrowser = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "file_type":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrunfile = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                            elif querytype == "white_hash":
                                csv_file = csv.writer(
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
                                    csv_file.writerow(tmp)
                                    firstrunhash = False
                                tmp = []
                                tmp.append(scope)
                                for key, value in data.items():
                                    tmp.append(value)
                                csv_file.writerow(tmp)

                    if cursor:
                        paramsnext = f"/web/api/{API_VERSION}/exclusions?{QUERY_LIMITS}&type={querytype}&countOnly=false&cursor={cursor}"
                        url = HOSTNAME.get() + paramsnext + exparam
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
            accounts = asyncio.create_task(get_accounts(session))
            await accounts

    async def runSites():
        async with aiohttp.ClientSession() as session:
            logger.debug("Running through sites")
            sites = asyncio.create_task(get_sites(session))
            await sites

    async def runGroups():
        async with aiohttp.ClientSession() as session:
            logger.debug("Running through groups")
            groups = asyncio.create_task(get_groups(session))
            await groups

    def getScope():
        logger.info("Getting user scope access")
        url = HOSTNAME.get() + f"/web/api/{API_VERSION}/user"
        r = requests.get(
            url,
            headers=headers,
            proxies={"http": PROXY.get(), "https": PROXY.get()},
        )
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
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
        logger.info("Getting account/site/group structure for %s", HOSTNAME.get())
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


def export_endpoint_tags():
    """Function to export Endpoint Tags from Console"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_ENDPOINT_TAGS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
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
        HOSTNAME.get()
        + f"/web/api/{API_VERSION}/agents/tags?includeChildren=true&includeParents=true&{QUERY_LIMITS}"
    )
    while url:
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": PROXY.get(), "https": PROXY.get()},
            verify=USE_SSL.get(),
        )
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
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
                paramsnext = f"/web/api/{API_VERSION}/agents/tags?includeChildren=true&includeParents=true&{QUERY_LIMITS}&cursor={cursor}"
                url = HOSTNAME.get() + paramsnext
                logger.debug("Next cursor found, updating URL: %s", url)
            else:
                logger.debug("No cursor found, setting URL to None")
                url = None
    logger.info("Done! Output file is - %s\n", export_csv)


def manage_endpoint_tags():
    """Add or Remove Endpoint Tags from Agents"""
    scroll_text = ScrolledText.ScrolledText(
        master=MANAGE_ENDPOINT_TAGS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    id_type = "computerName"
    if agent_id_type.get() == "uuid":
        id_type = "uuid"

    logger.debug("Specified an ID type of: %s", id_type)

    with open(INPUT_FILE.get()) as csv_file:
        logger.debug("Reading CSV: %s", INPUT_FILE.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        line_count = 0

        for row in csv_reader:
            logger.info("Updating Endpoint Tags for %s", row[0])
            url = HOSTNAME.get() + f"/web/api/{API_VERSION}/agents/actions/manage-tags"
            body = {
                "filter": {id_type: row[0]},
                "data": [
                    {
                        "operation": endpoint_tags_action.get(),
                        "tagId": tag_id_entry.get(),
                    }
                ],
            }
            logger.debug(
                "Calling API with the following:\nURL: %s\tData: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                json.dumps(body),
                headers,
                PROXY.get(),
                USE_SSL.get(),
            )
            response = requests.post(
                url,
                data=json.dumps(body),
                headers=headers,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
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
            logger.info("Finished! Input file %s was empty.", INPUT_FILE.get())
        else:
            logger.info("Finished! Processed %d lines.", line_count)


def export_local_config():
    """Export Agent Local Config"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_LOCAL_CONFIG_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    json_file = f"Local_Config_Export_{datestamp}.json"

    with open(INPUT_FILE.get()) as csv_file:
        logger.debug("Reading CSV: %s", INPUT_FILE.get())
        csv_reader = csv.reader(csv_file, delimiter=",")

        for row in csv_reader:
            logger.info("Getting Agent ID for Agent UUID: %s", row[0])
            url = HOSTNAME.get() + f"/web/api/{API_VERSION}/agents"
            agent_id = ""
            agent_config = ""
            param = {"uuid": row[0]}
            response = requests.get(
                url,
                params=param,
                headers=headers,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tParams: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                param,
                headers,
                PROXY.get(),
                USE_SSL.get(),
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
                HOSTNAME.get()
                + f"/web/api/{API_VERSION}/private/agents/{agent_id}/support-actions/configuration"
            )

            response = requests.get(
                url,
                params={},
                headers=headers,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tParams: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                param,
                headers,
                PROXY.get(),
                USE_SSL.get(),
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


def export_users():
    """Function to handle getting User Details and writing to CSV or XLSX"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_USERS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d_%f")
    logger.debug("User selected %s file type", user_output_type.get())
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
        HOSTNAME.get()
        + f"/web/api/{API_VERSION}/users?{QUERY_LIMITS}&sortOrder=asc&sortBy=email"
    )
    first_run = True
    total_users = 0
    users = {}

    logger.info("Getting Users list")
    while url:
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": PROXY.get(), "https": PROXY.get()},
            verify=USE_SSL.get(),
        )
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
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
                paramsnext = f"/web/api/{API_VERSION}/users?{QUERY_LIMITS}&sortOrder=asc&sortBy=email&cursor={cursor}"
                url = HOSTNAME.get() + paramsnext
                logger.debug("Next cursor found, updating URL: %s", url)
            else:
                logger.debug("No cursor found, setting URL to None")
                url = None

    if user_output_type.get() == "xlsx":
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
        "Done! Output file is - %s.%s\n", output_file_name, user_output_type.get()
    )


def export_ranger():
    """Function to handle exporting Ranger Inventory to CSV"""
    scroll_text = ScrolledText.ScrolledText(
        master=EXPORT_RANGER_INV_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    export_scope = export_ranger_scope.get()
    ranger_time_period = export_ranger_timeperiod.get()
    if export_scope == "sites":
        scope_param = "siteIds"
    else:
        scope_param = "accountIds"
    if not INPUT_FILE.get():
        logger.error("Must select a CSV containing Account or Site IDs")

    datestamp = datetime.datetime.now().strftime("%Y-%m-%d")

    logger.debug(
        "Input options:\n\tScope: %s\n\tScope ID CSV: %s\n\tTime Period: %s",
        export_scope,
        INPUT_FILE.get(),
        ranger_time_period,
    )
    with open(f"{str(INPUT_FILE.get())}") as csv_file:
        logger.debug("Reading CSV: %s", INPUT_FILE.get())
        csv_reader = csv.reader(csv_file, delimiter=",")
        for row in csv_reader:
            logger.info(
                "Exporting Ranger Inventory for %s scope ID: %s",
                export_scope.capitalize(),
                row[0],
            )
            firstrun = True
            endpoint = f"/web/api/{API_VERSION}/ranger/table-view?{QUERY_LIMITS}&period={ranger_time_period}&{scope_param}={row[0]}"
            url = HOSTNAME.get() + endpoint
            while url:
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
                )
                response = requests.get(
                    url,
                    headers=headers,
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
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
                        url = HOSTNAME.get() + paramsnext
                        logger.debug("Next cursor found, updating URL: %s", url)
                    else:
                        logger.debug("No cursor found, setting URL to None")
                        url = None
                logger.info("Finished writing to %s", csv_filename)
        logger.info("Done exporting Ranger Inventory.")


def export_account_ids():
    """Function to get all Account IDs for a tenant"""
    scroll_text = ScrolledText.ScrolledText(
        master=UPDATE_SYSTEM_CONFIG_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    endpoint = f"/web/api/{API_VERSION}/accounts?states=active&{QUERY_LIMITS}&sortBy=name&sortOrder=asc"
    url = HOSTNAME.get() + endpoint
    acct_ids_list = []

    while url:
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
        )
        response = requests.get(
            url,
            headers=headers,
            proxies={"http": PROXY.get(), "https": PROXY.get()},
            verify=USE_SSL.get(),
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
                for _, value in enumerate(data):
                    new_acct = {
                        "Account ID": value["id"],
                        "Account Name": value["name"],
                    }
                    acct_ids_list.append(new_acct)
            if cursor:
                paramsnext = endpoint + f"&cursor={cursor}"
                url = HOSTNAME.get() + paramsnext
                logger.debug("Next cursor found, updating URL: %s", url)
            else:
                logger.debug("No cursor found, setting URL to None")
                url = None

    csv_filename = "Account-IDs.csv"
    csv_columns = ["Account ID", "Account Name"]
    logger.debug("Opening %s to write", csv_filename)
    with open(csv_filename, "a+", newline="", encoding="utf-8") as file:
        csv_writer = csv.DictWriter(file, fieldnames=csv_columns)
        csv_writer.writeheader()
        logger.debug("Writing data to %s", csv_filename)

        for row in acct_ids_list:
            csv_writer.writerow(row)

        logger.info("Finished writing to %s", csv_filename)

    logger.info("Done exporting Account IDs.")


def bulk_resolve_threats():
    """Function to resolve multiple incidents by threat detail string search or SHA1"""
    scroll_text = ScrolledText.ScrolledText(
        master=BULK_RESOLVE_THREATS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    GET_LIMIT = 1
    POST_LIMIT = 2500  # Max per API Docs is 5000, in newer consoles
    RESOLVED_STATUS = "resolved"
    IS_RESOLVED = False
    THREAT_ENDPOINT = "/threats"
    NOTE_ENDPOINT = "/threats/notes"
    INCIDENT_ENDPOINT = "/threats/incident"
    PARTIAL_URL = f"{HOSTNAME.get()}/web/api/{API_VERSION}"
    site_ids = [x for x in site_ids_list.get().split(",")]
    new_verdict = selected_analyst_verdict.get()
    new_note = f"Analyst Verdict: '{new_verdict}'\nIncident Status: '{RESOLVED_STATUS}'\n\n- Set via S1 Manager."
    search_value = incident_search_value.get()
    multi_run = False
    rsession = requests.Session()

    logger.debug(
        "Creating appropriate params/payload for Incident Search Type: %s",
        incident_search_type.get(),
    )
    if incident_search_type.get() == "threat_name":
        get_params = {
            "limit": GET_LIMIT,
            "siteIds": site_ids,
            "resolved": IS_RESOLVED,
            "threatDetails__contains": f'"{search_value}"',
        }
        add_note_payload = json.dumps(
            {
                "filter": {
                    "limit": POST_LIMIT,
                    "siteIds": site_ids,
                    "resolved": IS_RESOLVED,
                    "threatDetails__contains": f'"{search_value}"',
                },
                "data": {"text": new_note},
            }
        )
        update_incident_payload = json.dumps(
            {
                "filter": {
                    "limit": POST_LIMIT,
                    "siteIds": site_ids,
                    "resolved": IS_RESOLVED,
                    "threatDetails__contains": f'"{search_value}"',
                },
                "data": {
                    "incidentStatus": RESOLVED_STATUS,
                    "analystVerdict": new_verdict,
                },
            }
        )
        logger.debug(
            "get_params = %s\nadd_note_payload = %s\nupdate_incident_payload = %s",
            get_params,
            add_note_payload,
            update_incident_payload,
        )
    else:
        get_params = {
            "limit": GET_LIMIT,
            "siteIds": site_ids,
            "resolved": IS_RESOLVED,
            "contentHashes": search_value,
        }
        add_note_payload = json.dumps(
            {
                "filter": {
                    "limit": POST_LIMIT,
                    "siteIds": site_ids,
                    "resolved": IS_RESOLVED,
                    "contentHashes": search_value,
                },
                "data": {"text": new_note},
            }
        )
        update_incident_payload = json.dumps(
            {
                "filter": {
                    "limit": POST_LIMIT,
                    "siteIds": site_ids,
                    "resolved": IS_RESOLVED,
                    "contentHashes": search_value,
                },
                "data": {
                    "incidentStatus": RESOLVED_STATUS,
                    "analystVerdict": new_verdict,
                },
            }
        )
        logger.debug(
            "get_params = %s\nadd_note_payload = %s\nupdate_incident_payload = %s",
            get_params,
            add_note_payload,
            update_incident_payload,
        )

    logger.info(
        "Checking for total number of unresolved incidents for: %s", search_value
    )

    with rsession as new_session:
        url = PARTIAL_URL + THREAT_ENDPOINT
        response = new_session.get(
            url=url,
            headers=headers,
            params=get_params,
            proxies={"http": PROXY.get(), "https": PROXY.get()},
            verify=USE_SSL.get(),
        )
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
        )
        if response.status_code != 200:
            logger.error(
                "Status: %s Problem with the request. Details - %s ",
                str(response.status_code),
                str(response.text),
            )
        response.raise_for_status()

        total_incidents = int(response.json()["pagination"]["totalItems"])

    if not total_incidents:
        logger.info(
            "Total unresolved incidents is %d. Nothing to change.",
            total_incidents,
        )
    else:
        logger.info(
            "Total unresolved incidents is %d. Starting to update and resolve incidents",
            total_incidents,
        )
        multi_run = True

    while multi_run:
        logger.info("Adding '%s' as a note to threat incidents", new_note)
        with rsession as new_session:
            url = PARTIAL_URL + NOTE_ENDPOINT
            response = new_session.post(
                url=url,
                headers=headers,
                data=add_note_payload,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
            )
            response.raise_for_status()

        logger.info(
            "Setting Analyst Verdict to '%s' and Incident Status to '%s'",
            new_verdict,
            RESOLVED_STATUS,
        )
        with rsession as new_session:
            url = PARTIAL_URL + INCIDENT_ENDPOINT
            response = new_session.post(
                url=url,
                headers=headers,
                data=update_incident_payload,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Status: %s Problem with the request. Details - %s ",
                    str(response.status_code),
                    str(response.text),
                )
            response.raise_for_status()

        logger.info("Checking if there are more incidents to update")
        with rsession as new_session:
            url = PARTIAL_URL + THREAT_ENDPOINT
            response = new_session.get(
                url=url,
                headers=headers,
                params=get_params,
                proxies={"http": PROXY.get(), "https": PROXY.get()},
                verify=USE_SSL.get(),
            )
            logger.debug(
                "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                url,
                headers,
                PROXY.get(),
                USE_SSL.get(),
            )
            if response.status_code != 200:
                logger.error(
                    "Status: %s Problem with the request. Details - %s ",
                    str(response.status_code),
                    str(response.text),
                )
            response.raise_for_status()

            total_incidents = int(response.json()["pagination"]["totalItems"])

        if not total_incidents:
            logger.info(
                "Total remaining unresolved incidents is '0', setting multi_run to False."
            )
            multi_run = False
        else:
            logger.info(
                "Total remaining unresolved incidents is %d. Continuing to update and resolve incidents",
                total_incidents,
            )

    logger.info("Done! Incidents resolved.\n")


def update_sys_config():
    """Function to read in a JSON configuration to update Account level system configuration settings."""
    scroll_text = ScrolledText.ScrolledText(
        master=UPDATE_SYSTEM_CONFIG_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    if not site_acct_ids_list.get():
        logger.error("Must input one or more Account IDs.")
    elif not INPUT_FILE.get():
        logger.error(
            "Must select a JSON file containing the new configuration to apply."
        )
    else:
        endpoint = f"/web/api/{API_VERSION}/system/configuration"
        url = HOSTNAME.get() + endpoint
        ids_list = [x for x in site_acct_ids_list.get().split(",")]
        logger.debug("ID List: %s", ids_list)
        file_name = Path(INPUT_FILE.get())
        valid_json = False

        with open(file_name, "r", encoding="utf-8") as file:
            logger.info("Reading %s", file_name)
            try:
                new_config = json.loads(file.read())
                logger.debug("%s appears to contain valid JSON", file_name.name)
                logger.debug("New config JSON contents: %s", new_config)
                valid_json = True
            except ValueError as exc:
                logger.error(
                    "%s possibly contains invalid JSON, please validate it and try again. %s",
                    file_name.name,
                    exc,
                )
                valid_json = False

        if valid_json:
            for new_id in ids_list:
                id_type = update_sites_or_accts.get()
                logger.debug("Current ID: %s", new_id)
                if isinstance(new_config, str):
                    new_config = json.loads(new_config)
                try:
                    logger.info(
                        "Updating JSON 'filter' with '%s':'%s'", id_type, new_id
                    )
                    if id_type == "siteIds":
                        new_config["filter"]["siteIds"] = new_id
                    elif id_type == "accountIds":
                        new_config["filter"]["accountIds"] = new_id
                except KeyError as err:
                    logger.error(
                        "Invalid key found in JSON. Ensure you selected the correct option between 'Sites' and 'Accounts', and that your JSON is correctly defined.\n%s",
                        err,
                    )
                    break

                new_config = json.dumps(new_config)
                logger.debug("Configuration: %s", new_config)

                response = requests.put(
                    url=url,
                    headers=headers,
                    data=new_config,
                    proxies={"http": PROXY.get(), "https": PROXY.get()},
                    verify=USE_SSL.get(),
                )
                logger.debug(
                    "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
                    url,
                    headers,
                    PROXY.get(),
                    USE_SSL.get(),
                )
                if response.status_code != 200:
                    logger.error(
                        "Status: %s Problem with the request. Details - %s ",
                        str(response.status_code),
                        str(response.text),
                    )
                    break
                response.raise_for_status()
                logger.info("System configuration updated for %s", new_id)

            logger.info("Finished.")


def bulk_enable_agents():
    scroll_text = ScrolledText.ScrolledText(
        master=BULK_ENABLE_AGENTS_FRAME, state="disabled", height=10
    )
    scroll_text.configure(font=ST_FONT)
    scroll_text.grid(row=13, column=0, columnspan=2, pady=10)
    text_handler = TextHandler(scroll_text)
    logging.basicConfig(
        filename=LOG_NAME,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
    )
    logger = logging.getLogger()
    logger.addHandler(text_handler)

    if not group_ids_list.get():
        logger.error("Must input one or more Group IDs.")
    else:
        endpoint = f"/web/api/{API_VERSION}/agents/actions/enable-agent"
        url = HOSTNAME.get() + endpoint
        group_ids = [x for x in group_ids_list.get().split(",")]
        logger.debug("ID List: %s", group_ids)

        payload = json.dumps(
            {
                "data": {
                    "shouldReboot": "false",
                },
                "filter": {"operationalStatesNin": "na", "groupIds": group_ids},
            }
        )

        response = requests.post(
            url=url,
            headers=headers,
            data=payload,
            proxies={"http": PROXY.get(), "https": PROXY.get()},
            verify=USE_SSL.get(),
        )
        logger.debug(
            "Calling API with the following:\nURL: %s\tHeaders: %s\tProxy: %s\tUse SSL: %s",
            url,
            headers,
            PROXY.get(),
            USE_SSL.get(),
        )
        if response.status_code != 200:
            logger.error(
                "Status: %s Problem with the request. Details - %s ",
                str(response.status_code),
                str(response.text),
            )
            response.raise_for_status()
        else:
            data = response.json()
            affected_agents = data.get("data").get("affected", "0")
            logger.info("Enable Agent action sent to %s", group_ids)
            logger.info("Total agents Enabled: %s", affected_agents)

        logger.info("Finished.")


# Login Menu Frame #############################
tk.Label(master=LOGIN_MENU_FRAME, image=LOGO).grid(
    row=0, column=0, columnspan=1, pady=20
)

console_address_label = tk.Label(
    master=LOGIN_MENU_FRAME,
    text="Management Console URL:",
)
console_address_label.grid(row=1, column=0, pady=2)

console_address_entry = ttk.Entry(master=LOGIN_MENU_FRAME, width=80)
console_address_entry.grid(row=2, column=0, pady=2)

api_token_label = tk.Label(master=LOGIN_MENU_FRAME, text="API Token:")
api_token_label.grid(row=3, column=0, pady=(10, 2))

api_token_entry = ttk.Entry(master=LOGIN_MENU_FRAME, width=80)
api_token_entry.grid(row=4, column=0, pady=2)

tk.Label(
    master=LOGIN_MENU_FRAME,
    text="*API Token provided must have sufficient permissions to perform a given action.",
    font=FRAME_SUBNOTE_FONT,
).grid(row=5, column=0, pady=5)

proxy_label = tk.Label(
    master=LOGIN_MENU_FRAME,
    text="Proxy (if required):",
)
proxy_label.grid(row=6, column=0, pady=(10, 2))

proxy_entry = ttk.Entry(master=LOGIN_MENU_FRAME, width=80)

proxy_entry.grid(row=7, column=0, pady=2)

use_ssl_switch = ttk.Checkbutton(
    master=LOGIN_MENU_FRAME,
    text="Use SSL",
    style="Switch",
    variable=USE_SSL,
    onvalue=True,
    offvalue=False,
)
use_ssl_switch.grid(row=8, column=0, pady=10)

login_button = ttk.Button(master=LOGIN_MENU_FRAME, text="Login", command=login)
login_button.grid(row=9, column=0, columnspan=2, ipady=5, pady=10)

if LOG_LEVEL == logging.DEBUG:
    ttk.Label(
        master=LOGIN_MENU_FRAME,
        text=f"S1 Manager launched with --debug. Be sure to delete {LOG_NAME} when finished.",
        font=FRAME_SUBNOTE_FONT,
        foreground=FRAME_NOTE_FG_COLOR,
    ).grid(row=10, column=0, pady=10, ipadx=5, ipady=5)

tk.Label(
    master=LOGIN_MENU_FRAME,
    text=f"SentinelOne API: {API_VERSION}\tS1 Manager: v{__version__}",
).grid(row=12, column=0, pady=(10, 5), sticky="s")
LOGIN_MENU_FRAME.pack()

# Main Menu Frame #############################
tk.Label(master=MAIN_MENU_FRAME, image=LOGO).grid(
    row=0, column=0, columnspan=4, pady=20
)

# Export - Column 0
ttk.Label(
    master=MAIN_MENU_FRAME, text="Export Operations", font=FRAME_SUBTITLE_FONT_UNDERLINE
).grid(row=1, column=0, columnspan=2, pady=20)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Deep Visiblity Events",
    command=partial(switch_frames, EXPORT_FROM_DV_FRAME),
    width=32,
).grid(row=2, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Activity Log",
    command=partial(switch_frames, EXPORT_ACTIVITY_LOG_FRAME),
    width=32,
).grid(row=3, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Endpoints",
    command=partial(switch_frames, EXPORT_ENDPOINTS_FRAME),
    width=32,
).grid(row=4, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Exclusions",
    command=partial(switch_frames, EXPORT_EXCLUSIONS_FRAME),
    width=32,
).grid(row=5, column=0, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Endpoint Tags",
    command=partial(switch_frames, EXPORT_ENDPOINT_TAGS_FRAME),
    width=32,
).grid(row=2, column=1, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Local Config",
    command=partial(switch_frames, EXPORT_LOCAL_CONFIG_FRAME),
    width=32,
).grid(row=3, column=1, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Users",
    command=partial(switch_frames, EXPORT_USERS_FRAME),
    width=32,
).grid(row=4, column=1, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Export Ranger Inventory",
    command=partial(switch_frames, EXPORT_RANGER_INV_FRAME),
    width=32,
).grid(row=5, column=1, sticky="ew", ipady=5, pady=5, padx=5)


# Manage - Column 2
tk.Label(
    master=MAIN_MENU_FRAME, text="Manage Operations", font=FRAME_SUBTITLE_FONT_UNDERLINE
).grid(row=1, column=2, columnspan=2, pady=20)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Upgrade Agents",
    command=partial(switch_frames, UPGRADE_FROM_CSV_FRAME),
    width=32,
).grid(row=2, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Move Agents",
    command=partial(switch_frames, MOVE_AGENTS_FRAME),
    width=32,
).grid(row=3, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Assign Customer Identifier",
    command=partial(switch_frames, ASSIGN_CUSTOMER_ID_FRAME),
    width=32,
).grid(row=4, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Decommission Agents",
    command=partial(switch_frames, DECOMMISSION_AGENTS_FRAME),
    width=32,
).grid(row=5, column=2, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Manage Endpoint Tags",
    command=partial(switch_frames, MANAGE_ENDPOINT_TAGS_FRAME),
    width=32,
).grid(row=2, column=3, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Bulk Resolve Threats",
    command=partial(switch_frames, BULK_RESOLVE_THREATS_FRAME),
    width=32,
).grid(row=3, column=3, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Bulk Enable Agents",
    command=partial(switch_frames, BULK_ENABLE_AGENTS_FRAME),
    width=32,
).grid(row=4, column=3, sticky="ew", ipady=5, pady=5, padx=5)
ttk.Button(
    master=MAIN_MENU_FRAME,
    text="Update System Config",
    command=partial(switch_frames, UPDATE_SYSTEM_CONFIG_FRAME),
    width=32,
).grid(row=5, column=3, sticky="ew", ipady=5, pady=5, padx=5)


if LOG_LEVEL == logging.DEBUG:
    ttk.Label(
        master=MAIN_MENU_FRAME,
        text=f"S1 Manager launched with --debug. Be sure to delete {LOG_NAME} when finished.",
        font=FRAME_SUBNOTE_FONT,
        foreground=FRAME_NOTE_FG_COLOR,
    ).grid(row=10, column=0, columnspan=4, pady=10, ipadx=5, ipady=5)

tk.Label(
    master=MAIN_MENU_FRAME,
    text="Note: Many of the processes can take a while to run. Be patient.",
    font=FRAME_SUBNOTE_FONT,
).grid(row=11, column=0, columnspan=4, padx=20, pady=20, sticky="s")


# Export from DV Frame #############################
tk.Label(
    master=EXPORT_FROM_DV_FRAME,
    text="Export Deep Visiblity Events",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=EXPORT_FROM_DV_FRAME,
    text="Export Deep Visibility events to an XLSX by query ID as reference",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=10)
tk.Label(master=EXPORT_FROM_DV_FRAME, text="1. Input Deep Visibility Query ID").grid(
    row=2, column=0, pady=2
)
query_id_entry = ttk.Entry(master=EXPORT_FROM_DV_FRAME, width=80)
query_id_entry.grid(row=3, column=0, pady=10)
ttk.Button(
    master=EXPORT_FROM_DV_FRAME,
    text="Export",
    command=export_from_dv,
).grid(row=4, column=0, pady=10)
ttk.Button(
    master=EXPORT_FROM_DV_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=5, column=0, ipadx=10, pady=10)


# Search and Export Activity Log Frame #############################
tk.Label(
    master=EXPORT_ACTIVITY_LOG_FRAME,
    text="Search and Export Activity Log",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=EXPORT_ACTIVITY_LOG_FRAME,
    text="Search Management Console Activity log and export results.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=10)
tk.Label(master=EXPORT_ACTIVITY_LOG_FRAME, text="1. Input FROM date (yyyy-mm-dd)").grid(
    row=2, column=0, pady=2
)
date_from = from_date_entry = ttk.Entry(master=EXPORT_ACTIVITY_LOG_FRAME, width=40)
from_date_entry.grid(row=3, column=0, pady=10)
tk.Label(master=EXPORT_ACTIVITY_LOG_FRAME, text="2. Input TO date (yyyy-mm-dd)").grid(
    row=4, column=0, pady=2
)
date_to = to_date_entry = ttk.Entry(master=EXPORT_ACTIVITY_LOG_FRAME, width=40)
to_date_entry.grid(row=5, column=0, pady=10)
tk.Label(master=EXPORT_ACTIVITY_LOG_FRAME, text="3. Input search string").grid(
    row=6, column=0, pady=2
)
string_search_entry = ttk.Entry(master=EXPORT_ACTIVITY_LOG_FRAME, width=80)
string_search_entry.grid(row=7, column=0, pady=2)
ttk.Button(
    master=EXPORT_ACTIVITY_LOG_FRAME,
    text="Search",
    command=partial(export_activity_log, True),
).grid(row=8, column=0, pady=10)
ttk.Button(
    master=EXPORT_ACTIVITY_LOG_FRAME,
    text="Export",
    command=partial(export_activity_log, False),
).grid(row=9, column=0, pady=10)
ttk.Button(
    master=EXPORT_ACTIVITY_LOG_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=10, column=0, ipadx=10, pady=10)


# Upgrade Agents Frame #############################
tk.Label(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Upgrade Agents",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Upgrade Agents to a specific package version by ID.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=UPGRADE_FROM_CSV_FRAME, text="1. Export Packages List to source Package ID"
).grid(row=2, column=0, padx=20, pady=2)
ttk.Button(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Export Packages List",
    command=partial(upgrade_from_csv, True),
).grid(row=3, column=0, pady=10)
tk.Label(master=UPGRADE_FROM_CSV_FRAME, text="2. Insert the Package ID").grid(
    row=4, column=0, pady=2
)
package_id_entry = ttk.Entry(master=UPGRADE_FROM_CSV_FRAME, width=80)
package_id_entry.grid(row=5, column=0, pady=2)
tk.Label(
    master=UPGRADE_FROM_CSV_FRAME,
    text="3. Select a CSV file containing a single column of endpoint names to upgrade",
).grid(row=6, column=0, padx=20, pady=2)
ttk.Button(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=7, column=0, pady=2)
tk.Label(master=UPGRADE_FROM_CSV_FRAME, textvariable=INPUT_FILE).grid(
    row=8, column=0, pady=2
)
use_schedule_switch = ttk.Checkbutton(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Use Schedule",
    style="Switch",
    variable=USE_SCHEDULE,
    onvalue=True,
    offvalue=False,
)
use_schedule_switch.grid(row=9, column=0, pady=10)
tk.Label(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Note: Will request upgrade immediately, unless 'Use Schedule' is toggled on.",
    font=FRAME_SUBNOTE_FONT,
).grid(row=10, column=0, pady=2)
ttk.Button(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Submit",
    command=partial(upgrade_from_csv, False),
).grid(row=11, column=0, pady=10)
ttk.Button(
    master=UPGRADE_FROM_CSV_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=12, column=0, ipadx=10, pady=10)


# Move Agents Frame #############################
tk.Label(
    master=MOVE_AGENTS_FRAME,
    text="Move Agents",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=MOVE_AGENTS_FRAME,
    text="Move Agents to specified Site ID and Group ID.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=MOVE_AGENTS_FRAME,
    text="If the target group is dynamic, the agent will be moved to the site only.",
).grid(row=2, column=0, pady=2)
tk.Label(master=MOVE_AGENTS_FRAME, text="1. Export Groups List to get group IDs").grid(
    row=3, column=0, pady=2
)
ttk.Button(
    master=MOVE_AGENTS_FRAME,
    text="Export Groups List",
    command=partial(move_agents, True),
).grid(row=4, column=0, pady=10)
tk.Label(
    master=MOVE_AGENTS_FRAME,
    text="2. Select a CSV file constructed of three columns:\nendpoints names, target group IDs, target site IDs",
).grid(row=5, column=0, padx=20, pady=10)
ttk.Button(master=MOVE_AGENTS_FRAME, text="Browse", command=select_csv_file).grid(
    row=6, column=0, pady=10
)
tk.Label(master=MOVE_AGENTS_FRAME, textvariable=INPUT_FILE).grid(
    row=7, column=0, pady=10
)
ttk.Button(
    master=MOVE_AGENTS_FRAME,
    text="Submit",
    command=partial(move_agents, False),
).grid(row=8, column=0, pady=10)
ttk.Button(
    master=MOVE_AGENTS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=9, column=0, ipadx=10, pady=10)


# Assign Customer Identifier Frame #############################
tk.Label(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="Assign Customer Identifier",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="Assign a Customer Identifier to one or more Agents.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, pady=2)
tk.Label(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="1. Input the Customer Identifier to assign",
).grid(row=2, column=0, padx=20, pady=2)
customer_id_entry = ttk.Entry(master=ASSIGN_CUSTOMER_ID_FRAME, width=80)
customer_id_entry.grid(row=3, column=0, pady=(2, 10))
tk.Label(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="2. Select a CSV file containing a single column with endpoint names",
).grid(row=4, column=0, padx=20, pady=2)
ttk.Button(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=5, column=0, pady=10)
tk.Label(master=ASSIGN_CUSTOMER_ID_FRAME, textvariable=INPUT_FILE).grid(
    row=6, column=0, pady=10
)
ttk.Button(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="Submit",
    command=assign_customer_id,
).grid(row=7, column=0, pady=10)
ttk.Button(
    master=ASSIGN_CUSTOMER_ID_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=8, column=0, ipadx=10, pady=10)


# Decommission Agents from CSV Frame #############################
tk.Label(
    master=DECOMMISSION_AGENTS_FRAME,
    text="Decommission Agents",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=DECOMMISSION_AGENTS_FRAME,
    text="1. Select a CSV file containing a single column of endpoint names to be decommissioned",
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=DECOMMISSION_AGENTS_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=2, column=0, pady=10)
tk.Label(master=DECOMMISSION_AGENTS_FRAME, textvariable=INPUT_FILE).grid(
    row=3, column=0, pady=10
)
ttk.Button(
    master=DECOMMISSION_AGENTS_FRAME,
    text="Submit",
    command=decommission_agents,
).grid(row=4, column=0, pady=10)
ttk.Button(
    master=DECOMMISSION_AGENTS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=5, column=0, ipadx=10, pady=10)


# Export all agents Frame #############################
tk.Label(
    master=EXPORT_ENDPOINTS_FRAME,
    text="Export Endpoints Light-Report",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=EXPORT_ENDPOINTS_FRAME,
    text="Exports up to 300,000 Agent details to a CSV, and converts it to XLSX",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
ttk.Button(
    master=EXPORT_ENDPOINTS_FRAME,
    text="Export",
    command=export_all_agents,
).grid(row=2, column=0, columnspan=2, pady=10)
ttk.Button(
    master=EXPORT_ENDPOINTS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=3, column=0, columnspan=2, ipadx=10, pady=10)


# Export Exclusions #############################
tk.Label(
    master=EXPORT_EXCLUSIONS_FRAME, text="Export Exclusions", font=FRAME_TITLE_FONT
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=EXPORT_EXCLUSIONS_FRAME,
    text="Exports all Exclusions to an XLSX",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=EXPORT_EXCLUSIONS_FRAME,
    text="Export",
    command=export_exclusions,
).grid(row=2, column=0, pady=10)
ttk.Button(
    master=EXPORT_EXCLUSIONS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=3, column=0, ipadx=10, pady=10)


# Export Endpoint Tag IDs Frame #############################
tk.Label(
    master=EXPORT_ENDPOINT_TAGS_FRAME,
    text="Export Endpoint Tags",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=EXPORT_ENDPOINT_TAGS_FRAME,
    text="Exports Endpoint Tag details to CSV for all scopes.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=2)
ttk.Button(
    master=EXPORT_ENDPOINT_TAGS_FRAME,
    text="Export",
    command=export_endpoint_tags,
).grid(row=2, column=0, pady=10)
ttk.Button(
    master=EXPORT_ENDPOINT_TAGS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=3, column=0, ipadx=10, pady=10)


# Manage Endpoint Tags Frame #############################
tk.Label(
    master=MANAGE_ENDPOINT_TAGS_FRAME,
    text="Manage Endpoint Tags",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=MANAGE_ENDPOINT_TAGS_FRAME,
    text="Add or Remove Endpoint Tags from Agents.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, pady=2)
tk.Label(master=MANAGE_ENDPOINT_TAGS_FRAME, text="1. Select Action").grid(
    row=2, column=0, columnspan=2, padx=20, pady=2
)
endpoint_tags_action = tk.StringVar()
endpoint_tags_action.set("add")
ttk.Radiobutton(
    MANAGE_ENDPOINT_TAGS_FRAME,
    text="Add Endpoint Tag",
    variable=endpoint_tags_action,
    value="add",
).grid(row=3, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    MANAGE_ENDPOINT_TAGS_FRAME,
    text="Remove Endpoint Tag",
    variable=endpoint_tags_action,
    value="remove",
).grid(row=3, column=1, padx=10, pady=2, sticky="w")
tk.Label(master=MANAGE_ENDPOINT_TAGS_FRAME, text="2. Input Endpoint Tag ID").grid(
    row=4, column=0, columnspan=2, padx=20, pady=2
)
tag_id_entry = ttk.Entry(master=MANAGE_ENDPOINT_TAGS_FRAME, width=80)
tag_id_entry.grid(row=5, column=0, columnspan=2, pady=(2, 10))
tk.Label(
    master=MANAGE_ENDPOINT_TAGS_FRAME,
    text="3. Select Agent Identifier type. This should align with your source CSV.",
).grid(row=6, column=0, columnspan=2, padx=20, pady=2)
agent_id_type = tk.StringVar()
agent_id_type.set("uuid")
ttk.Radiobutton(
    MANAGE_ENDPOINT_TAGS_FRAME, text="Agent UUID", variable=agent_id_type, value="uuid"
).grid(row=7, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    MANAGE_ENDPOINT_TAGS_FRAME,
    text="Endpoint Name",
    variable=agent_id_type,
    value="name",
).grid(row=7, column=1, padx=10, pady=2, sticky="w")
tk.Label(
    master=MANAGE_ENDPOINT_TAGS_FRAME,
    text="4. Select a CSV file containing a single column of values (uuids or endpoint names)",
).grid(row=8, column=0, columnspan=2, padx=20, pady=2)
ttk.Button(
    master=MANAGE_ENDPOINT_TAGS_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=9, column=0, columnspan=2, pady=10)
tk.Label(master=MANAGE_ENDPOINT_TAGS_FRAME, textvariable=INPUT_FILE).grid(
    row=10, column=0, columnspan=2, pady=10
)
ttk.Button(
    master=MANAGE_ENDPOINT_TAGS_FRAME, text="Submit", command=manage_endpoint_tags
).grid(row=11, column=0, columnspan=2, pady=10)
ttk.Button(
    master=MANAGE_ENDPOINT_TAGS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=12, column=0, columnspan=2, ipadx=10, pady=10)


# Export Agent Local Config Frame #############################
tk.Label(
    master=EXPORT_LOCAL_CONFIG_FRAME, text="Export Local Config", font=FRAME_TITLE_FONT
).grid(row=0, column=0, padx=20, pady=20)
tk.Label(
    master=EXPORT_LOCAL_CONFIG_FRAME,
    text="Exports the local agent configuration to a single JSON file.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, padx=20, pady=2)
tk.Label(
    master=EXPORT_LOCAL_CONFIG_FRAME,
    text="1. Select a CSV file containing a single column of agent UUIDs",
).grid(row=2, column=0, padx=20, pady=2)
ttk.Button(
    master=EXPORT_LOCAL_CONFIG_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=3, column=0, pady=10)
tk.Label(master=EXPORT_LOCAL_CONFIG_FRAME, textvariable=INPUT_FILE).grid(
    row=4, column=0, pady=10
)
ttk.Button(
    master=EXPORT_LOCAL_CONFIG_FRAME,
    text="Export",
    command=export_local_config,
).grid(row=5, column=0, pady=10)
ttk.Button(
    master=EXPORT_LOCAL_CONFIG_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=6, column=0, ipadx=10, pady=10)


# Export Users Frame #############################
tk.Label(master=EXPORT_USERS_FRAME, text="Export Users", font=FRAME_TITLE_FONT).grid(
    row=0, column=0, columnspan=2, padx=20, pady=20
)
tk.Label(
    master=EXPORT_USERS_FRAME,
    text="Exports User details to CSV.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
user_output_type = tk.StringVar()
user_output_type.set("csv")
ttk.Radiobutton(
    EXPORT_USERS_FRAME, text="CSV", variable=user_output_type, value="csv"
).grid(row=2, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    EXPORT_USERS_FRAME, text="XLSX", variable=user_output_type, value="xlsx"
).grid(row=2, column=1, padx=10, pady=2, sticky="w")
ttk.Button(
    master=EXPORT_USERS_FRAME,
    text="Export",
    command=export_users,
).grid(row=3, column=0, columnspan=2, pady=10)
ttk.Button(
    master=EXPORT_USERS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=4, column=0, columnspan=2, ipadx=10, pady=10)


# Export Ranger Inventory Frame #############################
tk.Label(
    master=EXPORT_RANGER_INV_FRAME,
    text="Export Ranger Inventory",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=EXPORT_RANGER_INV_FRAME,
    text="Exports Ranger Inventory details to CSV",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
tk.Label(
    master=EXPORT_RANGER_INV_FRAME,
    text="1. Select which scope type to export Ranger Inventory from.",
).grid(row=2, column=0, columnspan=2, padx=20, pady=2)
export_ranger_scope = tk.StringVar()
export_ranger_scope.set("accounts")
ttk.Radiobutton(
    EXPORT_RANGER_INV_FRAME,
    text="Account",
    variable=export_ranger_scope,
    value="accounts",
).grid(row=3, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    EXPORT_RANGER_INV_FRAME, text="Site", variable=export_ranger_scope, value="sites"
).grid(row=3, column=1, padx=10, pady=2, sticky="w")
tk.Label(
    master=EXPORT_RANGER_INV_FRAME,
    text="2. Select a CSV containing a single column of Account or Site IDs.",
).grid(row=4, column=0, columnspan=2, padx=20, pady=2)
ttk.Button(
    master=EXPORT_RANGER_INV_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=5, column=0, columnspan=2, pady=2)
tk.Label(master=EXPORT_RANGER_INV_FRAME, textvariable=INPUT_FILE).grid(
    row=6, column=0, columnspan=2, pady=2
)
tk.Label(
    master=EXPORT_RANGER_INV_FRAME,
    text="3. Specify time period for data export",
).grid(row=7, column=0, columnspan=2, padx=20, pady=2)
available_timeperiods = ("", "latest", "last12h", "last24h", "last3d", "last7d")
export_ranger_timeperiod = tk.StringVar()
export_ranger_timeperiod.set(available_timeperiods[1])
ttk.OptionMenu(
    EXPORT_RANGER_INV_FRAME, export_ranger_timeperiod, *available_timeperiods
).grid(row=8, column=0, columnspan=2, pady=10)
ttk.Button(
    master=EXPORT_RANGER_INV_FRAME,
    text="Export",
    command=export_ranger,
).grid(row=9, column=0, columnspan=2, pady=10)
ttk.Button(
    master=EXPORT_RANGER_INV_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=10, column=0, columnspan=2, ipadx=10, pady=10)


# Bulk Resolve Threats Frame #############################
tk.Label(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="Bulk Resolve Threats",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="Adds a note to each matching unresolved Incident then Resolves them with the specified Analyst Verdict.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
tk.Label(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="1. Select incident search type",
).grid(row=2, column=0, columnspan=2, padx=20, pady=2)
incident_search_type = tk.StringVar()
incident_search_type.set("threat_name")
ttk.Radiobutton(
    BULK_RESOLVE_THREATS_FRAME,
    text="Threat Name",
    variable=incident_search_type,
    value="threat_name",
).grid(row=3, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    BULK_RESOLVE_THREATS_FRAME,
    text="SHA1",
    variable=incident_search_type,
    value="content_hash",
).grid(row=3, column=1, padx=10, pady=2, sticky="w")
tk.Label(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="2. Input partial or complete threat name, or SHA1 based on above choice",
).grid(row=4, column=0, columnspan=2, padx=20, pady=2)
incident_search_value = ttk.Entry(master=BULK_RESOLVE_THREATS_FRAME, width=80)
incident_search_value.grid(row=5, column=0, columnspan=2, pady=10)
tk.Label(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="3. Select Analyst Verdict",
).grid(row=6, column=0, columnspan=2, padx=20, pady=2)
available_verdicts = ("", "undefined", "suspicious", "false_positive", "true_positive")
selected_analyst_verdict = tk.StringVar()
selected_analyst_verdict.set(available_verdicts[1])
ttk.OptionMenu(
    BULK_RESOLVE_THREATS_FRAME, selected_analyst_verdict, *available_verdicts
).grid(row=7, column=0, columnspan=2, pady=10)
tk.Label(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="4. Input one or more Site IDs, comma-separated with no spaces",
).grid(row=8, column=0, columnspan=2, padx=20, pady=2)
site_ids_list = ttk.Entry(master=BULK_RESOLVE_THREATS_FRAME, width=80)
site_ids_list.grid(row=9, column=0, columnspan=2, pady=10)
ttk.Button(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="Resolve Incidents",
    command=bulk_resolve_threats,
).grid(row=10, column=0, columnspan=2, pady=10)
ttk.Button(
    master=BULK_RESOLVE_THREATS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=11, column=0, columnspan=2, ipadx=10, pady=10)


# Update System Config Frame #############################
tk.Label(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="Update System Config",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="Updates the system configuration settings for one or more Accounts.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
tk.Label(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="1. Select if updating Sites or Accounts",
).grid(row=2, column=0, columnspan=2, padx=20, pady=2)
update_sites_or_accts = tk.StringVar()
update_sites_or_accts.set("siteIds")
ttk.Radiobutton(
    UPDATE_SYSTEM_CONFIG_FRAME,
    text="Sites",
    variable=update_sites_or_accts,
    value="siteIds",
).grid(row=3, column=0, padx=10, pady=2, sticky="e")
ttk.Radiobutton(
    UPDATE_SYSTEM_CONFIG_FRAME,
    text="Accounts",
    variable=update_sites_or_accts,
    value="accountIds",
).grid(row=3, column=1, padx=10, pady=2, sticky="w")
tk.Label(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="2. Input one or more IDs, comma-separated with no spaces",
).grid(row=4, column=0, columnspan=2, padx=20, pady=2)
site_acct_ids_list = ttk.Entry(master=UPDATE_SYSTEM_CONFIG_FRAME, width=80)
site_acct_ids_list.grid(row=5, column=0, columnspan=2, pady=10)
tk.Label(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="3. Select JSON file with new configuration",
).grid(row=6, column=0, columnspan=2, padx=20, pady=2)
ttk.Button(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="Browse",
    command=select_csv_file,
).grid(row=7, column=0, columnspan=2, pady=10)
tk.Label(master=UPDATE_SYSTEM_CONFIG_FRAME, textvariable=INPUT_FILE).grid(
    row=8, column=0, columnspan=2, pady=2
)
ttk.Button(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="Update",
    command=update_sys_config,
).grid(row=9, column=0, columnspan=2, pady=10)
ttk.Button(
    master=UPDATE_SYSTEM_CONFIG_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=10, column=0, columnspan=2, ipadx=10, pady=10)


# Bulk Enable Agents Frame #############################
tk.Label(
    master=BULK_ENABLE_AGENTS_FRAME,
    text="Bulk Enable Agents",
    font=FRAME_TITLE_FONT,
).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
tk.Label(
    master=BULK_ENABLE_AGENTS_FRAME,
    text="Send 'Enable Agent' action (without reboot) to all disabled agents in the specified list of Group IDs.",
    font=FRAME_SUBTITLE_FONT,
).grid(row=1, column=0, columnspan=2, padx=20, pady=2)
group_ids_list = ttk.Entry(master=BULK_ENABLE_AGENTS_FRAME, width=80)
group_ids_list.grid(row=2, column=0, columnspan=2, pady=10)
ttk.Button(
    master=BULK_ENABLE_AGENTS_FRAME,
    text="Enable",
    command=bulk_enable_agents,
).grid(row=3, column=0, columnspan=2, pady=10)
ttk.Button(
    master=BULK_ENABLE_AGENTS_FRAME,
    text="Back to Main Menu",
    command=go_back_to_mainpage,
).grid(row=4, column=0, columnspan=2, ipadx=10, pady=10)

window.mainloop()
