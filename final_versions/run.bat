@echo off
echo running rshb_write
python rshb_write.py
echo running psb_write
python psb_write.py
echo running sovkombank_write
python sovkombank_write.py
echo running mkb_write
python mkb_write.py
echo running vtb_write
python vtb_write.py
echo running gazprom_write
python gazprom_write.py
echo running sber_write
python sber_write.py
echo running alfa_write
python alfa_beta.py
pause

