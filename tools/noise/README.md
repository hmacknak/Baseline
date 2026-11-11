# Noise check

Runs the review checks over any folder of real `.xlsx` files and reports what fired, so false alarms can be
measured before a release. Nothing here ships in the add-in, and no workbooks are stored in this repo.

```
python3 tools/noise/xlsx2snap.py /tmp/corpus.json path/to/*.xlsx      # needs: pip install openpyxl
npx esbuild tools/noise/run.ts --bundle --platform=node --outfile=/tmp/noise.js
node /tmp/noise.js /tmp/corpus.json
```

Do not point it at client files on a shared machine: the snapshot file contains the cell contents.
