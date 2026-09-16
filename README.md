# DiscScan
It writes down the contents of your provided path into an Excel spread sheet.

USAGE: When window pops up write anything to wake it up. The first question asks you if you want to provide detailed path for scanning, if you write "y" you can paste your path. Otherwise you can just provide your disc's letter ("D" for example). Also its perormance speed depends heavily on system condition and file types you scan, media files take the longest because script uses a more advanced libraries to scan those.

Features on 1.0:
- It gives columns: index,	Nazwa pliku (file name),	Folder,	Pełna ścieżka (full path),	Rozmiar (file size in MB, rounded to one number after a comma),	Godzina utworzenia (creation time),	Data utworzenia (creation date),	Szerokość (width in px for pics and videos),	Wysokość (height in px for pics and videos)	Bitrate (kbps for music and videos)
- Works on detailed path or a whole disc
- Optimisation for different file types. It would just write simple details for non media files and when it detects the media file it uses media info to check media details. It decreases scan time drastically.
- Scanning speed is provided during the scan and you can see how many files it scanned already

Known issues:
- songs with album image also get width and height data in px, but you can sort them later in excel by their extension
- no completion time estimate, it would make the process take way longer
