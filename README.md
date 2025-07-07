# Spotify Playlist Exporter to Excel 🎵➡️📊

Easily export your Spotify playlists to an Excel (.xlsx) file with track details like name, artist, album, duration, and more!

---

## ✨ Features

- Export any public Spotify playlist to Excel (.xlsx)  
- Track info: Name, Artist(s), Album, Duration, Track URL  
- Playlist metadata: Playlist name, author, playlist URL  
- Auto-formats Excel for readability  
- Sums total playlist duration (hours:minutes)  
- CLI-friendly (uses environment variables or arguments)  
- Handles large playlists with pagination  

---

## Getting Started

### Prerequisites

- Python 3.7+  
- Spotify account  
- Spotify Developer account to create your own Spotify App (free)  

## Step 1: Clone the repository


git clone https://github.com/Mich0165/SpotifyToExcel.git
cd SpotifyToExcel

## Step 2: Create and activate a virtual environment (recommended)

python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate

## Step 3: Install dependencies

pip install -r requirements.txt

## Step 4: Register your Spotify Developer App
Spotify requires each user to have their own Spotify App credentials for authentication.

Visit the Spotify Developer Dashboard

Log in with your Spotify account

Click Create an App

Enter any App name and App description (examples below)

For Redirect URI, add exactly:

http://127.0.0.1:8888/callback/

Click Save

After creation, you'll see your Client ID and Client Secret on the app page

Example App Name and Description:
App Name: Spotify Playlist Exporter
Description: A simple app to export Spotify playlist data to Excel

## Step 5: Configure your environment variables
Create a .env file in the project root with the following content:

SPOTIPY_CLIENT_ID=your_client_id_here

SPOTIPY_CLIENT_SECRET=your_client_secret_here

SPOTIPY_REDIRECT_URI=http://127.0.0.1:8888/callback/

SPOTIPY_PLAYLIST_ID=your_playlist_id_here

Replace the placeholders with your actual Spotify app credentials.

Set SPOTIPY_PLAYLIST_ID to the playlist you want to export (playlist ID can be found in the Spotify app/share link).

## Step 6: Run the script
Run without arguments (it will use the playlist ID from .env):

python export_playlist.py
Or specify a playlist ID directly:

python export_playlist.py <playlist_id>

How to find a Playlist ID

Open Spotify desktop/web app

Navigate to the playlist you want to export

Click Share → Copy Spotify URI or Copy link

The URI looks like:

spotify:playlist:37i9dQZF1DXcBWIGoYBM5M
The part after the last colon is the playlist ID.

The link looks like:

https://open.spotify.com/playlist/37i9dQZF1DXcBWIGoYBM5M?si=...
The part after /playlist/ before the ? is the playlist ID.

---

Notes

Only public playlists and playlists you have access to can be exported.

If you encounter 403 Forbidden errors on audio features, the playlist or tracks might have access restrictions.

This script uses OAuth and the Spotify Web API, so you must have valid app credentials.

The redirect URI must match exactly what you entered in your Spotify developer app settings.

---

## 📜 License

This project is provided under the MIT License. Feel free to use, modify, and share it!


## 🙏 Credits

Built using Spotipy (Spotify Web API wrapper)
