set dDate=%date:~6%-%date:~3,2%-%date:~0,2%

git add .
git commit -m "%dDate%"
git branch -M main
git remote add origin https://github.com/mygithubnelsonr/MyExcelAddins.git
git push -u origin main

@echo last run: %Date% %Time%>> log.txt
