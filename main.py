#!/usr/bin/env python

# External dependencies needed to run this script:
#   1. FFmpeg (somewhere in PATH)

from datetime import datetime, timedelta
from os import chdir, mkdir
from subprocess import Popen
from sys import exit
from time import sleep

from playsound import playsound
from yt_dlp import YoutubeDL

# ---- GLOBALS ----

BELL = []  # bell schedule list
URLS = []  # list of URLs for media to download
OPTS = {  # yt-dlp arguments
    'format': 'mp3/bestaudio/best',
    'postprocessors': [{
        'key': 'FFmpegExtractAudio',
        'preferredcodec': 'mp3',
    }]
}


# ---- FUNCTIONS ----

def play_media(path):
    playsound(path, block=False)


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


def read_url_file():
    """
    Read list of URLs from "links.txt" and store them in the global "URLS" variable.
    """
    try:
        with open("../links.txt") as f:
            URLS.clear()  # clear URL list in case it already contains old data from previous day
            while line := f.readline().rstrip():
                URLS.append(line)
    except FileNotFoundError:
        print('"links.txt" does not exist, exiting now')
        exit(1)


def download_all():
    """
    Download all media from URLs in list with yt-dlp.
    """
    with YoutubeDL(OPTS) as ydl:
        extracted_urls = []
        for link in URLS:
            # extract the audio track URL from each link
            extracted_urls.append(ydl.extract_info(link, download=False)["url"])
    for i, link in enumerate(extracted_urls):
        # using FFmpeg:
        #   1. download all media at the same time
        #   2. convert first 1 minute to a mp3 file
        #   3. output file names are sequential
        Popen(f"ffmpeg -ss 00:00:00 -to 00:01:00 -i {link} -vn -ar 44100 -ac 2 -ab 192k -f mp3 bell_{i}.mp3")


def main():
    """
    Main program routine. Runs when script is executed.
    """
    setup_dirs()
    create_bell_schedule()
    read_url_file()
    download_all()


# ---- MAIN PROGRAM ----

if __name__ == '__main__':
    main()
