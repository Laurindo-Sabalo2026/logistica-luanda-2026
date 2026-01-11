name: Relatorio Diario Automatico

on:
  schedule:
    - cron: '0 7 * * *'
  workflow_dispatch:

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.9'
      - run: pip install pandas openpyxl fpdf matplotlib
      - env:
          MINHA_SENHA: ${{ secrets.MINHA_SENHA }}
        run: python main.py
