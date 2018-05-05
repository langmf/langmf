header
<?= if 1 then ?>
    kernel32.dll = <?= Print sys.ext.versiondll("kernel32.dll") ?>
    status = True
    msgbox True
<?= else ?>
    user32.dll = <?= Print sys.ext.versiondll("user32.dll") ?>
    status = False
    msgbox False
<?= end if ?>
footer