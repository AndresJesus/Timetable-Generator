import java.io.File;
import java.io.IOException;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class TimetableGenerator {
	public static GeneratingExcel g=new GeneratingExcel();
	
	public static void main(String[] args)throws IOException {		
		
		boolean imprimir=true; //	Crear el archivo excel.
		int minimo=0;
		int maximo=36;
		int longitud=35;
		int numeros[]=new int[longitud];	
			
//---------------------------------------------------------------------------------------------------------		
		
		// GENERANDO EL ARREGLO
	
		//	Voltea varianbles en caso de que estén al revez 
		if (maximo<minimo){
	            int aux=maximo;
	            maximo = minimo;
	            minimo=aux;
	            }

			
		if( (maximo-minimo) >= (longitud-1)){ //	Comprueba si la longitud es permitida
				
	            int numero_elementos=0;
	            
	            boolean encontrado;
	            int aleatorio;
	            
	            
	            while(numero_elementos<longitud){
			
	                aleatorio=generaNumeroAlAleatorio(minimo, maximo);
	                encontrado=false;
	                
	                for(int i=0;i<numeros.length && !encontrado ;i++){
	                    if(aleatorio==numeros[i] ){
	                        encontrado=true;
	                    }
	                  
	                }
	                
	                if(!encontrado){
	                	numeros[numero_elementos++] = aleatorio;
	                }
	                
	                
	            }
	
			}else{	System.out.println("No existe ese numero"); }			

//-------------------------------------------------------------------------------------------------
		
		String[] materias=new String[35];
		
		materias[0]="educ.fisica ";
		materias[1]="educ.fisica ";
		
		materias[2]="ciencia y tecnologia ";
		materias[3]="ciencia y tecnologia ";
		materias[4]="ciencia y tecnologia ";
		
		materias[5]="estetica ";
		materias[6]="estetica ";

		materias[7]="matematica ";
		materias[8]="matematica ";
		materias[9]="matematica ";
		materias[10]="matematica ";
		materias[11]="matematica ";
		materias[12]="matematica ";
		
		materias[13]="atencion ";

		materias[14]="educ. para el trabajo ";
		materias[15]="educ. para el trabajo ";
		
		materias[16]="folklore ";
		
		materias[17]="ingles ";
		materias[18]="ingles ";
		
		materias[19]="sociales ";
		materias[20]="sociales ";
		materias[21]="sociales ";
		
		materias[22]="biblioteca/computacion ";
		materias[23]="biblioteca/computacion ";
		materias[24]="biblioteca/computacion ";
		
		materias[25]="lengua ";
		materias[26]="lengua ";
		materias[27]="lengua ";
		materias[28]="lengua ";
		materias[29]="lengua ";
		
		materias[30]="recreo ";
		materias[31]="recreo ";
		materias[32]="recreo ";
		materias[33]="recreo ";
		materias[34]="recreo ";
		
		
			String [][] entrada= new String[7][5];
			int contador=-1;
			
			
			//Recorriendo el  excel e imprimiendolo
			for(int i=0;i<7;i++){
				
				System.out.println("\n");
				
				for(int j=0;j<5;j++){
					
					contador++;

					numeros[contador]--;
							
					entrada[i][j]=materias[numeros[contador]];
					System.out.print(entrada[i][j]);

					if(imprimir){
						String ruta="/Users/PC1/Desktop/salida.xls";
						g.generarExcel(entrada, ruta);
						}	
					
					}
				}	

			
			//	Muestra el combinacion de numeros que se usó para generar el arreglo
			System.out.println("Mostrar arreglo");
			
            for(int i=0;i<numeros.length;i++){
                
                System.out.print(numeros[i]+" ");

                }	
	
	
	}
			
		
			
					
			
	//	Funcion que genera numeros aleatorios dentro del parametro de Máximo y Mímino
	
	public static int generaNumeroAlAleatorio(int minimo, int maximo){
		
	    int num=(int)Math.abs(Math.floor(Math.random()*(minimo-(maximo+1))+(maximo)));
	    return num;
	}
}

class GeneratingExcel {

	public void generarExcel(String[][] entrada, String ruta) {

		try {
			WorkbookSettings conf = new WorkbookSettings();
			conf.setEncoding("ISO-8859-1");
			WritableWorkbook woorkBook = Workbook.createWorkbook(new File(ruta), conf);

			WritableSheet sheet = woorkBook.createSheet("Resultado", 0);

			WritableFont h = new WritableFont(WritableFont.TAHOMA, 10, WritableFont.NO_BOLD);
			WritableCellFormat hformat = new WritableCellFormat(h);

			CellView cell = new CellView();

			//hformat.setAlignment(Alignment.CENTRE);

			for (int x = 0; x < 5; x++) {
				cell = sheet.getColumnView(x);
				cell.setAutosize(true);
				sheet.setColumnView(x, cell);

			}



			for (int i = 0; i < 7; i++) {
				for (int j = 0; j < 5; j++) {


					try {
						sheet.addCell(new jxl.write.Label(j, i, entrada[i][j], hformat));
					} catch (RowsExceededException e) {

						e.printStackTrace();
					} catch (WriteException e) {

						e.printStackTrace();
					}


				}

			}

			woorkBook.write();

			try {
				woorkBook.close();
			} catch (WriteException e) { e.printStackTrace(); }



		} catch (IOException ex) {}

	}

}