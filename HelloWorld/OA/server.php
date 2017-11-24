<?php
session_start();
$num = $_POST['num'];

$_SESSION['num_1'] = "1234";

echo  "<script>window.location.href='device_add_list.php'</script>";