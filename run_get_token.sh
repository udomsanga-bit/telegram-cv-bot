#!/bin/bash
cd ~/Downloads/cv_bot
pip3 install google-auth-oauthlib --break-system-packages -q 2>/dev/null || pip3 install google-auth-oauthlib -q
python3 get_token.py
