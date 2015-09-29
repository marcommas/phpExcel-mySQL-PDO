<!DOCTYPE HTML>

<?php
/*
 * Cria conexão com o Banco
 */
require_once('lib/l_conexao.php');
$conn = db_conPDO();


	if(isset($_POST['geraExcel']))
	{
		/** Error reporting */
		error_reporting(E_ALL);
		ini_set('display_errors', TRUE);
		ini_set('display_startup_errors', TRUE);
		date_default_timezone_set('Europe/London');

		if (PHP_SAPI == 'cli')
			die('This example should only be run from a Web Browser');

		/** Include PHPExcel */
		require_once dirname(__FILE__) . '/lib/PHPExcel/Classes/PHPExcel.php';


		// Create new PHPExcel object
		$objPHPExcel = new PHPExcel();

		//Propriedades do documento
		$objPHPExcel->getProperties()->setCreator("Criador")
									 ->setLastModifiedBy("Última modificação")
									 ->setTitle("Título Relatório")
									 ->setSubject("Subtítulo Relatório")
									 ->setDescription("Descrição")
									 ->setKeywords("office 2007 openxml php")
									 ->setCategory("Relatório");

		try{
			$sql = " SELECT titulo, dados"; 
			$sql .= " FROM tabela ";	
			$sql .= " WHERE dados=:dados";

			//PREPARA A SQL
			$stmt = $conn->prepare($sql);

            $stmt->bindParam(':dados', dados, PDO::PARAM_INT);
			
			$stmt->execute();

			if($stmt->rowCount() > 0)
			{
				$i=0;
				while($row = $stmt->fetch(PDO::FETCH_OBJ)) {
					$cont=1;
                    $titulo = $row->titulo;	
					$dados = $row->dados;		
				
			    	// Criando uma nova planilha dentro do arquivo
					$objPHPExcel->createSheet();

					// Agora, vamos adicionar os dados na planinha da posição $i
					$objPHPExcel->setActiveSheetIndex($i);

					//TIRA A LINHA DE GRADE
					$objPHPExcel->getActiveSheet()->setShowGridlines(false);

					//DEFINE AS LARGURAS DAS COLUNAS
					$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
					$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
				

					//UNE AS COLUNAS
					$objPHPExcel->getActiveSheet()->mergeCells("A1:B1");

					// Define o título da planilha, seria o nome da aba
					$objPHPExcel->getActiveSheet()->setTitle($titulo);
                    
                    //ALINHAMENTO CENTRAL
					$objPHPExcel->getActiveSheet()->getStyle('A1:A2')->getAlignment()
													                    	->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                	$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true)
										                               		->setSize(16);

					//MOSTRA O NOME DA ENQUETE NO INÍCIO DO EXCEL
					$objPHPExcel->getActiveSheet()->SetCellValue('A1', $dados);

					$i++;
				}
			}
		}catch(PDOException $e) {
			echo 'ERROR: ' . $e->getMessage();
		}

		
		// Define a planilha como ativa sendo a primeira, assim quando abrir o arquivo será a que virá aberta como padrão
		$objPHPExcel->setActiveSheetIndex(0);
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');

		// Redirect output to a client’s web browser (Excel5)
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        header('Content-Disposition: attachment;filename="NomeDoArquivo.xlsx"');

		header('Cache-Control: max-age=0');
		// If you're serving to IE 9, then the following may be needed
		header('Cache-Control: max-age=1');

		// If you're serving to IE over SSL, then the following may be needed
		header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
		header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		header ('Pragma: public'); // HTTP/1.0

		
		ob_end_clean();
		$objWriter->save('php://output');
		exit;

	}	
	
?>

<html lang="pt-br">
    <head>
	    <meta charset="UTF-8" />
	</head>

	<body>
        <form id="formulario" name="formulario" action="relatorio.php" method="post" enctype="multipart/form-data" >
            <div class='btn-group'>
                <input type='submit' name='geraExcel' id='geraExcel' value='Gerar Excel' />
            </div>
        </form>                  
	</body>
	
</html>


