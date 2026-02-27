import os
os.chdir('/home/ubuntu/RA_tool')
os.system("git reset --hard")
os.system("git pull")
os.system ('sudo chown :www-data ~/RA_tool/')
os.system('sudo chmod 777 -R  ~/RA_tool/')
os.system ('sudo systemctl restart apache2')
os.system ('cat /etc/apache2/envvars')
