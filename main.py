#!/usr/bin/env python

import fnmatch
import json
import logging
import pickle
from datetime import datetime, timedelta
from itertools import cycle
from os import chdir, mkdir, listdir
from os.path import isfile
from random import shuffle
from subprocess import Popen, run
from sys import exit
from time import sleep

import coloredlogs
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException as openpyxl_InvalidFileException
from yt_dlp import YoutubeDL


# ---- CLASSES ----

class DummyLogger:
    # This class is needed to disable the output from yt-dlp
    def debug(self, msg):
        pass

    def info(self, msg):
        pass

    def warning(self, msg):
        pass

    def error(self, msg):
        pass


# ---- FUNCTIONS ----

def check_env():
    """
    Checks if all needed external dependencies are available. Exits script if any are missing.
    """
    # checks for ffmpeg and ffplay
    if not isfile("ffmpeg.exe") or not isfile("ffplay.exe"):
        logging.critical('Missing "ffmpeg" or "ffplay" (or both). '
                         'If you are using Windows, run "setup.bat" to download both.')
        input("\nPress ENTER to exit")
        exit(1)


def show_intro():
    """
    Displays the intro message to the terminal at the start of the script.
    """
    message = '''
        ---- Usage Instructions ----

Press "Ctrl + C" to play the next bell immediately.
Press "Ctrl + C" to stop playing the bell immediately.
Close the terminal window to stop this script at any time.
'''
    print(message)


def show_version():
    """
    Outputs the script's version to the logger.
    """
    try:
        with open("VERSION", mode='r') as f:
            logging.info(f"aus-bell version {f.read(7)}")
    except FileNotFoundError:
        logging.error('"VERSION" file is missing, continuing anyway')


def show_bell_schedule():
    """
    Outputs the bell schedule to the logger.
    """
    logging.debug("Bell schedule:")
    logging.debug("--------------")
    for time in BELL_SCHEDULE:
        logging.debug(time.strftime('%I:%M %p'))


def play_media(file_path):
    """
    Plays audio files using "ffplay". Can be stopped early with "Ctrl+C".
    :param file_path: Path to the audio file
    """
    try:
        run(f"../ffplay -nodisp -autoexit -loglevel fatal {file_path}")
    except KeyboardInterrupt:
        logging.warning(f"User requested to stop the bell early at {datetime.now().strftime('%I:%M %p')}")


def set_play_order():
    """
    Sets the order that the bell will play.
    """
    # for each file in media folder...
    for file in fnmatch.filter(listdir(), 'bell_*.mkv'):
        PLAYLIST.append(file)
    shuffle(PLAYLIST)
    global PLAY_CYCLE
    PLAY_CYCLE = cycle(PLAYLIST)
    logging.info(f"playlist with {len(PLAYLIST)} songs has been shuffled")


def ring_bell():
    """
    Rings the next bell in the playlist.
    """
    song = next(PLAY_CYCLE)
    logging.info(f"playing file: {song}")
    play_media(song)


def sleep_until(target: datetime):
    """
    Sleeps script until the target date and time. Pressing "Ctrl+C" will cancel the sleep.
    :param target: date and time to sleep until (must be a "datetime" object)
    """
    now = datetime.now()
    delta = target - now

    if delta > timedelta(0):
        logging.info(f"Waiting until next bell at {target.strftime('%I:%M %p')}")
        try:
            sleep(delta.total_seconds())
        except KeyboardInterrupt:
            logging.warning(f"User requested to play the bell early at {datetime.now().strftime('%I:%M %p')}")
    else:
        raise ValueError('"sleep_until()" cannot sleep a negative amount of time.')


def create_bell_schedule():
    """
    Initiates a list with the current day's bell schedule.
    """
    today = datetime.now().replace(second=0, microsecond=0)
    for time in CFG_DICT["bell_schedule"]:
        BELL_SCHEDULE.append(today.replace(hour=time['h'], minute=time['m']))


def setup_dirs():
    """
    Prepare needed directories and change working directory to the "media" folder.
    """
    try:
        mkdir("media")
        logging.info('"media" directory was created')
    except FileExistsError:
        pass
    chdir("media")


def setup_paths():
    """
    Prepare needed file paths.
    """
    global LINKS_PATH, LOG_PATH, CACHE_PATH
    try:
        LINKS_PATH = CFG_DICT["links_spreadsheet_path"]
        logging.info(f'Path to links spreadsheet: "{LINKS_PATH}"')
    except KeyError:
        logging.critical('"config.json" file does not contain the key "links_spreadsheet_path". Stopping now.')
        input("\nPress ENTER to exit")
        exit(1)
    try:
        LOG_PATH = CFG_DICT["log_file_path"]
        logging.info(f'Path to log file: "{LOG_PATH}"')
    except KeyError:
        logging.critical('"config.json" file does not contain the key "log_file_path". Stopping now.')
        input("\nPress ENTER to exit")
        exit(1)
    CACHE_PATH = "../cache.dat"  # this path is static


def set_current_media_list():
    """
    Fills out the global "CURRENT_MEDIA_LIST" with "bell numbers".
    """
    for file in fnmatch.filter(listdir(), 'bell_*.mkv'):
        CURRENT_MEDIA_LIST.append(int(file[5]))


def setup_logging():
    """
    Sets up the config for the logger.
    """
    # initiate file logger
    logging.basicConfig(format='%(asctime)s %(levelname)s: %(message)s',
                        level=logging.DEBUG,
                        filename=LOG_PATH,
                        filemode='w'
                        )

    # initiate terminal logger (in color)
    coloredlogs.install(level="DEBUG")
    logging.info("Script started, logging initiated")


def load_prev_urls():
    """
    Loads the URL list from the last time this script was ran.
    This is needed to compare if changes/deletions to the spreadsheet were made.
    """
    global PREV_URLS
    try:
        with open(CACHE_PATH, "rb") as f:
            PREV_URLS = pickle.load(f)
        logging.info("Loaded list of URLs from the previous run")
    except FileNotFoundError:
        # continuing without a cache file is ok, but show warning
        logging.warning("No cache file found, cannot load previous list of URLs")


def save_curr_urls():
    """
    Saves the current URL list to a cache file.
    Used in conjunction with "load_prev_urls()"
    """
    with open(CACHE_PATH, "wb") as f:
        pickle.dump(ALL_URLS, f, pickle.HIGHEST_PROTOCOL)


def read_url_file():
    """
    Read list of URLs from "links.xlsx" spreadsheet and store them in the global "URLS" variable.
    Also sets up the global "NEEDED_MEDIA_LIST" list.
    """
    try:
        wb = load_workbook(filename=LINKS_PATH, read_only=True)
        ws = wb["Sheet1"]
        for count, row in enumerate(ws.iter_rows(max_col=1, values_only=True)):
            link = row[0]
            ALL_URLS.append(link)
            if count not in CURRENT_MEDIA_LIST:
                if link is not None:
                    URLS.append(link)
                    NEEDED_MEDIA_LIST.append(count)
    except FileNotFoundError:
        logging.critical(f'"{LINKS_PATH}" does not exist, stopping now')
        input("\nPress ENTER to exit")
        exit(1)
    except PermissionError:
        logging.critical(f'Permission was denied to read spreadsheet file (located at: "{LINKS_PATH}"). '
                         'The file is likely open in Excel. Close Excel and run this script again. Stopping now.')
        input("\nPress ENTER to exit")
        exit(1)
    except openpyxl_InvalidFileException:
        logging.critical("Cannot not open links spreadsheet, it is an invalid file type. "
                         "Supported formats are: .xlsx, .xlsm, .xltx, .xltm.")
        input("\nPress ENTER to exit")
        exit(1)


def download_all():
    """
    Downloads all media from URLs in list with yt-dlp,
    converts first "MEDIA_LENGTH" minute(s) to a mkv file (copies original codec)
    """
    with YoutubeDL(OPTS) as ydl:
        extracted_urls = []
        logging.info(f"extracting audio URLs from {len(URLS)} links with yt-dlp")
        for link in URLS:
            # extract the audio track URL from each link
            try:
                extracted_urls.append(ydl.extract_info(link, download=False)["url"])
                logging.info(f"yt-dlp extracted the audio URL for {link}")
            except (TypeError, KeyError):
                # tried to download an invalid link, just skip this one
                logging.warning(f"yt-dlp skipped an invalid URL: {link}")
                extracted_urls.append(None)
                continue
    download_count = 0
    for link in extracted_urls:
        if link is not None:
            # count all valid links
            download_count += 1
    logging.info(f"Downloading {download_count} files with ffmpeg")
    ffmpeg_processes = []
    for file_num, link in zip(NEEDED_MEDIA_LIST, extracted_urls):
        if link is None:
            # skip invalid link
            continue
        # using ffmpeg:
        #   1. download all media at the same time
        #   2. convert first "MEDIA_LENGTH" minute(s) to a mkv file (copy audio codec)
        #   3. output file names are sequential (they match the order in spreadsheet)
        #   4. wait to return until all ffmpeg instances are finished
        ffmpeg_processes.append(
            Popen(
                f"../ffmpeg -loglevel fatal -n -ss 00:00:00 -to 00:{MEDIA_LENGTH} -i {link} "
                f"-vn -c:a copy bell_{file_num}.mkv"
            )
        )
    for process in ffmpeg_processes:
        # wait for all downloads to finish
        while process.poll() is None:
            sleep(0.5)
        else:
            logging.info(f"ffmpeg PID {process.pid} is finished")


def load_config():
    """
    Loads all data from the config file "config.json".
    """
    global CFG_DICT
    try:
        with open("../config.json", 'r') as f:
            CFG_DICT = json.load(f)
    except FileNotFoundError:
        logging.critical(
            'file "config.json" does not exist in root directory of program. Cannot load config. Stopping now.'
        )
        input("\nPress ENTER to exit")
        exit(1)


def main():
    """
    Main program routine. Runs when script is executed.
    """
    show_intro()
    setup_logging()
    show_version()
    check_env()
    setup_dirs()
    load_config()
    setup_paths()
    set_current_media_list()
    create_bell_schedule()
    show_bell_schedule()
    load_prev_urls()
    read_url_file()
    save_curr_urls()
    download_all()
    set_play_order()
    for time in BELL_SCHEDULE:
        try:
            sleep_until(time)
        except ValueError:
            # bell has already happened, so skip it
            continue
        ring_bell()
    logging.debug("Script finished cleanly")
    input("\nPress ENTER to exit")


# ---- GLOBALS ----

# all global variables needed for "main()"
MEDIA_LENGTH = "03:00"  # max length to trim downloaded media in MM:SS format (string)
CURRENT_MEDIA_LIST = []  # list of numbers representing available media files
NEEDED_MEDIA_LIST = []  # list of numbers representing the needed media files
BELL_SCHEDULE = []  # bell schedule list
URLS = []  # list of URLs to media to be downloaded
PREV_URLS = []  # list of URLS from the previous run
ALL_URLS = []  # all URLs that are in the spreadsheet (valid or not!)
LINKS_PATH = None  # file path to "links.xlsx" (where media URLs are stored)
LOG_PATH = ""  # file path to "bell.log" (where the logger saves to)
CACHE_PATH = ""  # file path to "cache.dat" (where the previous run's URLs are stored)
PLAYLIST = []  # bell play order
PLAY_CYCLE = cycle([])  # circular list version of PLAYLIST
CFG_DICT = {}  # dictionary of data read from "config.json"
OPTS = {  # yt-dlp arguments
    'format': 'bestaudio/best',
    'ignoreerrors': True,
    "quiet": True,
    "noprogress": True,
    "no_warnings": True,
    "logger": DummyLogger()
}

if __name__ == '__main__':
    main()
