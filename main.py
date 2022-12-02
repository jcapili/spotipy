from pytube import YouTube
from moviepy.editor import *

import os, subprocess, eyed3

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of the spreadsheet.
SPREADSHEET_ID = '1IVtVncMGfUO9JA4_5lEh_JYn24sEGmHXmSganCUoAIo'
SHEET_ID = 0
RANGE_NAME = 'Main!A2:F'

DOWNLOAD_PATH = os.path.expanduser('~/Downloads')

def parse_deletion_ranges(indeces):
    """
    Calculates the continues ranges of indeces given a list of 
    individual indeces. This function also takes into account
    the fact that Google will carry out each deletion request
    individually, so the indeces on each subsequent deletion
    request will change.

    @param indeces List of integers
    @return List of lists of integers, where each list contains
        only 2 integers denoting the start and end of the ranges
    """
    if len(indeces) == 0:
        return []

    ranges = []
    current_range = []
    for i in indeces:
        if current_range == []:
            current_range = [i] * 2

        if i == current_range[1]:
            continue
        elif i == current_range[1] + 1:
            current_range[1] = i
        else:
            # Need to add 1 because delete function is [inclusive, exclusive]
            current_range[1] += 1
            ranges.append(current_range)
            current_range = [i] * 2

    current_range[1] += 1
    ranges.append(current_range)

    # Adjust ranges to account for new indeces after each successful deletion
    total_deleted = 0
    for r in ranges:
        r[0] -= total_deleted
        r[1] -= total_deleted
        total_deleted += r[1] - r[0]

    return ranges

def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=RANGE_NAME).execute()
        values = result.get('values', [])
        
        successful_indeces = []

        # Iterate over the list backwards to maintain sheet order in iTunes
        for i in range(len(values) - 1, -1, -1):
            [link, title, artist, album, genre] = values[i]
            try:
                # Extract audio file as .mp4 and download
                yt = YouTube(link)
                mp4 = yt.streams.filter(only_audio=True).first()
                mp4_out_file = mp4.download(output_path=DOWNLOAD_PATH)

                # Write the video file as an .mp3 to allow it to open
                # in iTunes
                audio = AudioFileClip(mp4_out_file)
                audio_out_file = f'{DOWNLOAD_PATH}/{title}.mp3'
                audio.write_audiofile(audio_out_file, bitrate='128k')

                # Add metadata for iTunes
                metadata_audio = eyed3.load(audio_out_file)
                metadata_audio.tag.artist = artist
                metadata_audio.tag.album = album
                metadata_audio.tag.genre = genre
                metadata_audio.tag.save()

                # Play the .mp3 file to import into iTunes, then pause and delete
                subprocess.call(['open', audio_out_file])
                subprocess.call(['osascript', '-e', 'tell application "iTunes" to pause'])
                for file in [mp4_out_file, audio_out_file]:
                    os.remove(file)

                successful_indeces.append(i + 1)  # Don't delete the first row (0th index)
            except Exception as e:
                print(f'Error for "{title}": {e}')
                continue

        # Delete successful rows from sheets doc
        requests = []
        for _range in parse_deletion_ranges(sorted(successful_indeces)):
            requests.append({
                'deleteDimension': {
                    'range': {
                        'sheetId': SHEET_ID,
                        'dimension': 'ROWS',
                        'start_index': _range[0],
                        'end_index': _range[1],
                    }
                }
            })

        if requests != []:
            sheet.batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={
                    'requests': requests
                }
            ).execute()
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()