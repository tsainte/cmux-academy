# cmux Academy

A self-paced course for learning cmux. Open `index.html` to start.

| File | What it is |
| --- | --- |
| `index.html` | Hub with the curriculum and your progress |
| `slides.html` | reveal.js deck, five modules (needs a network connection for the CDN) |
| `tasks.html` | 37 small hands-on tasks with a checklist (progress saved in the browser) |
| `drill.html` | Flashcards for shortcuts and CLI commands; missed cards come back more often |
| `reference.html` | Searchable shortcuts and commands |
| `reading.html` | Official docs, ordered by module |
| `cmux-academy.xlsx` | Tracker with Progress, Tasks, Shortcuts, CLI and Reading list tabs |

## Run it inside cmux

```sh
cd ~/Developer/study/terminal/cmux-academy
python3 -m http.server 8000
# in another pane
cmux browser open http://localhost:8000
```

## Change the content

Lessons, tasks, shortcuts, commands and links live in `tools/build.py`. After editing:

```sh
python3 tools/build.py            # regenerates assets/data.js
# the .xlsx also needs openpyxl:  pip install openpyxl
```

The slides are hand-written in `slides.html`.

Content comes from cmux.com/docs and `cmux --help` (cmux 0.64.x). Shortcuts can change between releases.
