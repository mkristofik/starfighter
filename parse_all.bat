rem Convert all the new file format ships to JSON. ~dpn translates to drive,
rem path, base filename. Quoting the paths is necessary because the filenames
rem contain spaces.
for %%f in (ships\*.sw2) do (
	python parse_starfighter.py "%%f" > "%%~dpnf.json"
)

rem And convert all the old file format ships to JSON too.
for %%f in ("old ships"\*.sws) do (
	python parse_starfighter.py "%%f" > "%%~dpnf.json"
)
