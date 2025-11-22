#!/bin/bash
cd ~/Robo/Robo
source venv/bin/activate
python3 mensagens_diaria.py >> ~/Robo/Robo/logs.txt 2>&1

