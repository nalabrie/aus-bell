#!/usr/bin/env python

import pickle
from datetime import datetime, timedelta
from os import chdir, mkdir, path, getenv
from subprocess import Popen
from sys import exit
from time import sleep

from openpyxl import load_workbook
from playsound import playsound
from yt_dlp import YoutubeDL

# External dependencies needed to run this script:
#   1. ffmpeg (somewhere in PATH)

# ---- GLOBALS ----

COUNT = 0  # counter for naming downloaded media files
BELL = []  # bell schedule list
URLS = []  # list of URLs to media to be downloaded
PREV_URLS = []  # list of URLs that were used the previous time the URL list was loaded from "links.xlsx"
LINKS_PATH = None  # file path to "links.xlsx"
OPTS = {  # yt-dlp arguments
    'format': 'mp3/bestaudio/best',
    'ignoreerrors': True,
    'postprocessors': [{
        'key': 'FFmpegExtractAudio',
        'preferredcodec': 'mp3',
    }]
}


# ---- FUNCTIONS ----

def save_cache():
    """
    Save the state of needed variables (using Pickle) for the next time this script is ran.
    """
    state = (COUNT, PREV_URLS)
    with open("../cache.dat", "wb") as f:
        pickle.dump(state, f, pickle.HIGHEST_PROTOCOL)


def load_cache():
    """
    Load the cache from the Pickle file. If there is no cache file, do nothing.
    """
    try:
        with open("../cache.dat", "rb") as f:
            state = pickle.load(f)
    except FileNotFoundError:
        print("no cache file found, skipping cache load")
        return
    global COUNT, PREV_URLS
    COUNT = state[0]
    PREV_URLS = state[1]


def play_media(file_path):
    playsound(file_path, block=False)


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
    BELL.clear()  # clear bell schedule in case it already contains old data from previous day
    today = datetime.now().replace(second=0, microsecond=0)
    BELL.append(today.replace(hour=9, minute=15))
    BELL.append(today.replace(hour=10, minute=12))
    BELL.append(today.replace(hour=11, minute=15))
    BELL.append(today.replace(hour=12, minute=12))
    BELL.append(today.replace(hour=13, minute=42))
    BELL.append(today.replace(hour=14, minute=42))
    BELL.append(today.replace(hour=15, minute=40))


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
    global LINKS_PATH
    home = getenv("USERPROFILE")
    LINKS_PATH = path.join(home, "OneDrive", "OneDrive - ausohio.com", "bell", "links.xlsx")


def read_url_file():
    """
    Read list of URLs from "links.xlsx" spreadsheet and store them in the global "URLS" variable.
    """
    try:
        URLS.clear()  # clear URL list in case it already contains old data from previous day
        wb = load_workbook(filename=LINKS_PATH)
        ws = wb["Sheet1"]
        for row in ws.iter_rows(max_col=1, values_only=True):
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
    updates previous URL list (PREV_URLS).
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
    global COUNT, PREV_URLS
    PREV_URLS.clear()
    PREV_URLS = URLS.copy()  # URLs are done being extracted, so copy list to previous URL list
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
                f"ffmpeg -loglevel quiet -n -ss 00:00:00 -to 00:01:00 -i {link} -vn -ar 44100 -ac 2 -ab 192k -f mp3 bell_{COUNT}.mp3"))
        COUNT += 1
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
    load_cache()
    create_bell_schedule()
    read_url_file()
    download_all()
    save_cache()


# ---- MAIN PROGRAM ----

if __name__ == '__main__':
    main()
