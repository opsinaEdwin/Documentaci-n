<?php

// es la importacion de una libreria que gestiona la creacion de Excel.
require "vendor/autoload.php";
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$dbhost = "localhost";
$dbname = "basedatos";
$dbchar = "utf8mb4";
$dbuser = "root";
$dbpass = "";

// la conexion puede ser por PDO o por mysqli_conenct
$pdo = new PDO(

"mysql:host=$dbhost;charset=$dbchar;dbname=$dbname",$dbuser, $dbpass, [
  PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
  PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_NAMED]
);

// vamos a crear un excel vacio, una plantilla sin ningun dato


$spreadsheet = new Spreadsheet();
// se va a escribir por defecto en la primera sheet
$sheet = $spreadsheet->getActiveSheet();
// se le da un titulo
$sheet->setTitle("Usuarios");


// insertar los datos de la base de datos al sheet que se creo y que esta vacio

//se prepara la consulta
$stmt = $pdo->prepare("SELECT * FROM `usuario`");

// se ejecuta la consulta
$stmt->execute();

// variable controladora
$i = 1;


// insertar los datos de la base de datos al sheet que se creo y que esta vacio
while ($row = $stmt->fetch()) 
{

   	$sheet->setCellValue("A".$i, $row["idusuario"]);
	$sheet->setCellValue("B".$i, $row["nombre"]);
	$sheet->setCellValue("C".$i, $row["apellido"]);
	$sheet->setCellValue("D".$i, $row["correo"]);
	$sheet->setCellValue("E".$i, $row["contraseña"]);
	$i++;
}
$writer = new Xlsx($spreadsheet);
$writer->save("usuarios.xlsx");
?>


