# busy-spider
A simple crawler to gather mail of local buisnesses via google. The spider doens't obey bot.txt rules. I made it just for educational purposes!

## Requirements

Python 3.12. Runtime packages are pinned in `requirements.txt`.

## Setup

```bash
python -m pip install -r requirements-dev.txt
python -m pytest
```

`python main.py` and `python to_excel.py` behave as before. `to_excel.py` reads `last_call.csv` and writes `hyperlink.xlsx`.
