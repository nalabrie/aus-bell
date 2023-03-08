#!/usr/bin/env python

import fnmatch
import logging
from datetime import datetime, timedelta
from os import chdir, mkdir, path, getenv, listdir
from random import shuffle
from subprocess import Popen
from sys import exit
from time import sleep

from openpyxl import load_workbook
from playsound import playsound
from yt_dlp import YoutubeDL

# External dependencies needed to run this script:
#   1. ffmpeg (somewhere in PATH)

# ---- GLOBALS ----

MEDIA_FILE_COUNT = 0  # counter for naming downloaded media files
BELL_SCHEDULE = []  # bell schedule list
URLS = []  # list of URLs to media to be downloaded
LINKS_PATH = None  # file path to "links.xlsx" (where media URLs are stored)
LOG_PATH = None  # file path to "bell.log" (where the logger saves to)
PLAYLIST = []  # bell play order
OPTS = {  # yt-dlp arguments
    'format': 'mp3/bestaudio/best',
    'ignoreerrors': True,
    'postprocessors': [{
        'key': 'FFmpegExtractAudio',
        'preferredcodec': 'mp3',
    }]
}


# ---- FUNCTIONS ----

def play_media(file_path):
    playsound(file_path, block=False)


def set_play_order():
    """
    Sets the order that the bell will play.
    """
    # for each file in media folder...
    for file in fnmatch.filter(listdir(), 'bell_*.mp3'):
        PLAYLIST.append(file)
    shuffle(PLAYLIST)


def ring_bell():
    song = PLAYLIST.pop()
    print(f"playing file: {song}")
    play_media(song)


def sleep_until(target: datetime):
    """
    Sleeps script until the target date and time.

    :param target: date and time to sleep until (must be a "datetime" object)
    """
    now = datetime.now()
    delta = target - now

    if delta > timedelta(0):
        print(f"Sleeping script until {target.ctime()}")
        sleep(delta.total_seconds())
    else:
        raise ValueError('"sleep_until()" cannot sleep a negative amount of time.')


def create_bell_schedule():
    BELL_SCHEDULE.clear()  # clear bell schedule in case it already contains old data from previous day
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


def set_count():
    global MEDIA_FILE_COUNT
    for _ in fnmatch.filter(listdir(), 'bell_*.mp3'):
        MEDIA_FILE_COUNT += 1


def setup_logging():
    """
    Sets up the config for the logger.
    """
    logging.basicConfig(format='%(asctime)s %(message)s', filename=LOG_PATH, level=logging.DEBUG)


def read_url_file():
    """
    Read list of URLs from "links.xlsx" spreadsheet and store them in the global "URLS" variable.
    """
    try:
        URLS.clear()  # clear URL list in case it already contains old data from previous day
        wb = load_workbook(filename=LINKS_PATH)
        ws = wb["Sheet1"]
        for row in ws.iter_rows(min_row=MEDIA_FILE_COUNT + 1, max_col=1, values_only=True):
            link = row[0]
            if link is not None:
                URLS.append(link)
    except FileNotFoundError:
        print(f'"{LINKS_PATH}" does not exist, exiting now')
        exit(1)


def download_all():
    """
    Downloads all media from URLs in list with yt-dlp,
    converts first 1 minute to mp3,
    """
    with YoutubeDL(OPTS) as ydl:
        extracted_urls = []
        for link in URLS:
            # extract the audio track URL from each link
            try:
                extracted_urls.append(ydl.extract_info(link, download=False)["url"])
            except TypeError:
                # tried to download an invalid link, just skip this one
                pass
    global MEDIA_FILE_COUNT
    ffmpeg_processes = []
    print(f"Downloading {len(extracted_urls)} files with ffmpeg")
    for link in extracted_urls:
        # using ffmpeg:
        #   1. download all media at the same time
        #   2. convert first 1 minute to a mp3 file
        #   3. output file names are sequential (sequence is saved and can resume next run)
        #   4. wait to return until all ffmpeg instances are finished
        ffmpeg_processes.append(
            Popen(
                f"ffmpeg -loglevel quiet -n -ss 00:00:00 -to 00:01:00 -i {link} -vn -ar 44100 -ac 2 -ab 192k -f mp3 bell_{MEDIA_FILE_COUNT}.mp3"))
        MEDIA_FILE_COUNT += 1
    for process in ffmpeg_processes:
        # wait for all downloads to finish
        while process.poll() is None:
            sleep(0.5)
        else:
            print(f"ffmpeg PID {process.pid} is finished")


def main():
    """
    Main program routine. Runs when script is executed.
    """
    setup_dirs()
    setup_paths()
    setup_logging()
    logging.info("Script started")
    set_count()
    create_bell_schedule()
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


# ---- MAIN PROGRAM ----

if __name__ == '__main__':
    main()
