
<?php
/******************************************************************
@<autor>Comassetto</autor>
@<data>19/12/2012</data>
@<programa>l_conexao.php</programa>

@<objetivo>
@Executar a conexão com o banco de dados
@</objetivo>
@<instrucoes>
@</instrucoes>
******************************************************************/
function db_conPDO(){
	
	$host = 'localhost';
	$db = 'banco';
	$user = 'root';
	$pass = 'password';
	
	try
	{
	    $conn = new PDO( "mysql:host=$host; dbname=$db", $user, $pass);
	    return $conn;
	}
	catch ( PDOException $e )
	{
		echo 'Erro ao conectar com o Banco de Dados: ' . $e->getMessage();
	}	

}

?>