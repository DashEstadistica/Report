<?php
header('Access-Control-Allow-Origin: *');
if (isset($_GET['msg'])) {
    file_put_contents('browser_errors.txt', $_GET['msg'] . "\n", FILE_APPEND);
    echo "OK";
}
?>
