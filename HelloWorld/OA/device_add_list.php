<?php
session_start();
// store session data
$_SESSION['views']=1;

echo "this is ".$_SESSION['views'];
?>

<form action="server.php" method="post">
    <input type="text" name="num" value="<? echo $_SESSION['views'] ?>">
    <input type="submit" value="OK">
</form>

<h1><? $_SESSION['views'] ?></h1>