#!/usr/bin/env python

import fnmatch
import logging
from datetime import datetime, timedelta
from itertools import cycle
from os import chdir, mkdir, path, getenv, listdir
from os.path import isfile
from random import shuffle
from subprocess import Popen, run
from sys import exit
from time import sleep

import coloredlogs
from openpyxl import load_workbook
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
    BELL_SCHEDULE.append(today.replace(hour=9, minute=15))
    BELL_SCHEDULE.append(today.replace(hour=10, minute=12))
    BELL_SCHEDULE.append(today.replace(hour=11, minute=15))
    BELL_SCHEDULE.append(today.replace(hour=12, minute=12))
    BELL_SCHEDULE.append(today.replace(hour=13, minute=42))
    BELL_SCHEDULE.append(today.replace(hour=14, minute=42))
    BELL_SCHEDULE.append(today.replace(hour=15, minute=40))


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
    global LINKS_PATH, LOG_PATH
    home = getenv("USERPROFILE")
    LINKS_PATH = path.join(home, "OneDrive", "OneDrive - ausohio.com", "bell", "links.xlsx")
    LOG_PATH = path.join(home, "OneDrive", "OneDrive - ausohio.com", "bell", "bell.log")
    logging.info(f'Path to links spreadsheet: "{LINKS_PATH}"')
    logging.info(f'Path to log file: "{LOG_PATH}"')


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


def read_url_file():
    """
    Read list of URLs from "links.xlsx" spreadsheet and store them in the global "URLS" variable.
    Also sets up the global "NEEDED_MEDIA_LIST" list.
    """
    try:
        wb = load_workbook(filename=LINKS_PATH, read_only=True)
        ws = wb["Sheet1"]
        for count, row in enumerate(ws.iter_rows(max_col=1, values_only=True)):
            if count not in CURRENT_MEDIA_LIST:
                link = row[0]
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
            except TypeError:
                # tried to download an invalid link, just skip this one
                logging.warning(f"yt-dlp skipped an invalid URL: {link}")
                extracted_urls.append(None)
                continue
    ffmpeg_processes = []
    logging.info(f"Downloading {len(extracted_urls)} files with ffmpeg")
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


def main():
    """
    Main program routine. Runs when script is executed.
    """
    show_intro()
    setup_logging()
    show_version()
    check_env()
    setup_dirs()
    setup_paths()
    set_current_media_list()
    create_bell_schedule()
    show_bell_schedule()
    read_url_file()
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
LINKS_PATH = None  # file path to "links.xlsx" (where media URLs are stored)
LOG_PATH = ""  # file path to "bell.log" (where the logger saves to)
PLAYLIST = []  # bell play order
PLAY_CYCLE = cycle([])  # circular list version of PLAYLIST
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
