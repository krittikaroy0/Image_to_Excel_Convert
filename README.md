1. To install Composer globally on Windows:
Download and run Composer-Setup.exe.

During setup:

Select your PHP path (typically C:\xampp\php\php.exe if using XAMPP).

It will automatically configure composer in your system PATH.
2. Check Composer: composer --version

3. Enable ext-gd in PHP
If using XAMPP:
Go to your XAMPP installation folder:

makefile
C:\xampp\php
Open:
php.ini
Search for:
php.ini
;extension=gd
Remove the ; to uncomment:

extension=gd
Save the file.

Restart Apache:
Open XAMPP Control Panel.
Stop Apache.
Start Apache again.

4. Verify GD is enabled
Run in git bash
php -m

✅ You should see:
gd

5. in the list of modules.
Or create a phpinfo.php with:
<?php phpinfo(); ?>

6. Retry Composer Installation
Go to your project folder:
cd /path/to/your/project

Run:
composer require phpoffice/phpspreadsheet
