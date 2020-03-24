import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


public class compareexcel_step7 
{

	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception 
	{
	
		
		String ActFolderParentPath="E:/Automation/pav/PavRegression/excelcomparision";
		String ActFileFolderPath=ActFolderParentPath + File.separator + "Actual";
		String ExpFileFolderPath=ActFolderParentPath + File.separator + "Expected";
		String ActFilePath=null;
		String ExpFilePath=null;
		
		/*String DBUsername=args[0];
		String DBPassword=args[1];
		String DBHostname=args[2];
		String DBServName=args[3];
		String aPlanCd=args[4];
		String aCycleDate=args[5];*/
		
		//Read the Parameters
		File fXmlFile = new File("E:/Automation/pav/PavRegression/InputParameters.xml");
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(fXmlFile);
		doc.getDocumentElement().normalize();
		NodeList nList = doc.getElementsByTagName("DownStreamValues");
		Node nNode=nList.item(0);
		Element eElement = (Element) nNode;
		String bCycleDate=eElement.getElementsByTagName("bCycleDate").item(0).getTextContent();
		String DBHostname=eElement.getElementsByTagName("DBHostname").item(0).getTextContent();
		String DBServName=eElement.getElementsByTagName("DBServName").item(0).getTextContent();
		String DBUsername=eElement.getElementsByTagName("DBUsername").item(0).getTextContent();
		String DBPassword=eElement.getElementsByTagName("DBPassword").item(0).getTextContent();
		String aplancode=eElement.getElementsByTagName("aplancode").item(0).getTextContent();
		
		System.out.println("bCycleDate:[" + bCycleDate + "] " + "DBHostname:[" + DBHostname + "] "  + "DBServName:[" + DBServName + "] "  + "DBUsername:[" + DBUsername + "] " + "DBPassword:[" + DBPassword + "] "  + "aplancode:[" + aplancode + "] ");
        
				
		/*	DBUsername= System.getenv("DBUsername");		
		DBHostname= System.getenv("DBHostname");
		DBServName= System.getenv("DBServName");
		aPlanCd= System.getenv("aPlanCd");
		aCycleDate= System.getenv("aCycleDate");  */  

		/*		String DBUsername="aaa3205";
		String DBPassword="VaApr2017123";
		String DBHostname="florasit-scan";
		String DBServName="FEPLSIT_RPTW";
		String aPlanCd="082";
		String aCycleDate="12/22/2016";*/
		
		/*SimpleDateFormat mStandardDateTimeFormat = new SimpleDateFormat("MM/dd/yyyy");
		Date CurDate=new Date();
		Calendar aCalen = Calendar.getInstance();
		aCalen.setTime(CurDate); 
		aCalen.add(Calendar.DATE, -1);
		String aCycleDate = mStandardDateTimeFormat.format(aCalen.getTime());*/
		

		ActFilePath=ActFileFolderPath + File.separator + "PlanCode_" + aplancode + "_GL.xlsx" ;
		
		GetActualFile(DBUsername, DBPassword,DBHostname, DBServName, ActFilePath, aplancode,bCycleDate);

	 	File ActFilespath = new File(ActFileFolderPath);
	    File [] files = ActFilespath.listFiles();
	    for (int i = 0; i < files.length; i++)
	    {
	        if (files[i].isFile())
	        { 
	    		 boolean resfailflag=false;
	        	 
	        	 ExpFilePath=ExpFileFolderPath + File.separator + files[i].getName();
	        	
	             FileInputStream ActFile = new FileInputStream(new File(ActFilePath));	
	             File aExpFile=new File(ExpFilePath);
		         if(!aExpFile.exists())
		         {
		        	 	resfailflag=true;
		        	 	System.out.println("Expected File Doesn't Exists:[" + ExpFilePath + "]"); 
		         }else
		         {
	        	   	 FileInputStream ExpFile = new FileInputStream(new File(ExpFilePath));
		    		 XSSFWorkbook ExpWorkbook = new XSSFWorkbook(ExpFilePath);
		             XSSFWorkbook ActWorkbook = new XSSFWorkbook(ActFilePath);
		    		
		             XSSFSheet ExpSheet = ExpWorkbook.getSheetAt(0);
		             XSSFSheet ActSheet = ActWorkbook.getSheetAt(0);
	
		          	StringBuilder aFinalData = new StringBuilder();
		    	    aFinalData.append("<html><head><title>Test Result </title></head>");
		    	    aFinalData.append("<body>");
		    	    aFinalData.append("<table style=max-width:100px; border=\"1\" bordercolor=\"#000000\">");
		    	    
		     		if(!compareExptoActSheets(ExpSheet, ActSheet, aFinalData)) 
		            {
		     			resfailflag=true;
		            }
		     		
		     		if(!compareActtoExpSheets(ExpSheet, ActSheet, aFinalData)) 
		            {
		     			resfailflag=true;
		            }
		             
		            ExpFile.close();
		            ActFile.close();
	
		            aFinalData.append("</table></body></html>");
		            String aFilepath=WriteToFile(aFinalData.toString(),"pavglresult");
		            //String aFilepath=WriteToFile(aFinalData.toString(),files[i].getName().substring(0, files[i].getName().length()-5) + "_Result");
		            
		            if(resfailflag)
		    			System.out.println("ERROR: Excel Sheets with name " + files[i].getName().substring(0, files[i].getName().length()-5) + " are not Equal.Result Report Generated under:" + aFilepath);
		            	
		       }
	        }
	    }
	}	
	
	public static String WriteToFile(String fileContent, String fileName) throws IOException 
	{
	    String projectPath = System.getProperty("user.dir");
	    
		System.out.println("projectPath:" + projectPath);
		
       // String tempFile = projectPath + File.separator + fileName + ".html";
        String tempFile = "E:/Automation/pav/PavRegression/excelcomparision/Results/" + fileName + ".html";
    	
        System.out.println("fileName:" + fileName);
        System.out.println("tempFile:" + tempFile);
        
        File file = new File(tempFile);
        if (file.exists()) 
        {
            try {
                File newFileName = new File(projectPath + File.separator+ "backup_"+fileName);
                file.renameTo(newFileName);
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        
        //write to file with OutputStreamWriter
        OutputStream outputStream = new FileOutputStream(file.getAbsoluteFile());
        Writer writer=new OutputStreamWriter(outputStream);
        writer.write(fileContent);
        writer.close();
        return projectPath;
    }
		
	
	public static boolean compareExptoActSheets(XSSFSheet ExpSheet, XSSFSheet ActSheet,StringBuilder aFinalData) throws IOException 
	{
        int ExpFirstRow = ExpSheet.getFirstRowNum(); 
        int ExplastRow = ExpSheet.getLastRowNum();
        boolean equalSheets = true;
        
        int ActFirstRow = ActSheet.getFirstRowNum(); 
        int ActlastRow = ActSheet.getLastRowNum();

        //Append Column Names to Final String builder
        XSSFRow Firstrow = ExpSheet.getRow(0);
        aFinalData.append("<tr>");
        for(int m=Firstrow.getFirstCellNum(); m <= Firstrow.getLastCellNum()-1; m++) 
        {
        	if(m==Firstrow.getFirstCellNum())
        		aFinalData.append("<td></td>");
        	
            XSSFCell Colcell = Firstrow.getCell(m);
            String Colcellval=Colcell.toString();
            aFinalData.append("<td style= \"background-color:silver\">" + Colcellval + "</td>");
        }
        aFinalData.append("</tr>");
        
        //Loop through Expected sheet rows and search in Actual sheet
        for(int i=ExpFirstRow; i <= ExplastRow; i++) 
        {
            boolean rowexists=false;
            XSSFRow Exprow = ExpSheet.getRow(i);
            
            XSSFCell ClmNumExpcell = Exprow.getCell(0);
            String ExpClmNum=ClmNumExpcell.toString();
            
            XSSFCell AdjResExpcell = Exprow.getCell(5);
            String ExpAdjRes=AdjResExpcell.toString();
            
            String ExpChkStr=ExpClmNum + ":" + ExpAdjRes;
            
            //Loop through Actual sheet
            for(int j=ActFirstRow; j <= ActlastRow; j++) 
            {
                XSSFRow Actrow = ActSheet.getRow(j);
                XSSFCell ClmNumActcell = Actrow.getCell(0);
                String ActClmNum=ClmNumActcell.toString();
                
                XSSFCell AdjResActcell = Actrow.getCell(5);
                String ActAdjRes=AdjResActcell.toString();
                
                String ActChkStr=ActClmNum + ":" + ActAdjRes;
                
                if(ExpChkStr.equalsIgnoreCase(ActChkStr))
                {
                	rowexists=true;
                	if(!compareTwoRows(Exprow, Actrow, aFinalData)) 
                    {
                        equalSheets = false;
                    }
                	break;
                }
            }
            
            if(!rowexists)
            {
            	  int firstExpcell = Exprow.getFirstCellNum();
                  int lastExpcell = Exprow.getLastCellNum();
                  
                  // Compare all cells in a row
                  aFinalData.append("<tr>");
                  aFinalData.append("<td style= \"background-color:red\">" + "Row Missing in Actual:" + " </td>");
                  String cellval=null;
                  for(int l=firstExpcell; l <= lastExpcell-1; l++) 
                  {
                      XSSFCell Expcell = Exprow.getCell(l);
                      if(Exprow.getCell(l) != null)
                      {
                    	  cellval=Expcell.toString();
                      }
                      aFinalData.append("<td>" + cellval + "</td>");
                  }
                  aFinalData.append("</tr>");
                  equalSheets = false;
            }
        }
        return equalSheets;
    }
	
	public static boolean compareActtoExpSheets(XSSFSheet ExpSheet, XSSFSheet ActSheet,StringBuilder aFinalData) throws IOException 
	{
        int ExpFirstRow = ExpSheet.getFirstRowNum(); 
        int ExplastRow = ExpSheet.getLastRowNum();
        boolean equalSheets = true;

        int ActFirstRow = ActSheet.getFirstRowNum(); 
        int ActlastRow = ActSheet.getLastRowNum();
     
        //Loop through Actual sheet rows and search in Expected sheet
        for(int i=ActFirstRow; i <= ActlastRow; i++) 
        {
            boolean rowexists=false;
            XSSFRow Actrow = ActSheet.getRow(i);
            
            XSSFCell ClmNumActcell = Actrow.getCell(0);
            String ActClmNum=ClmNumActcell.toString();
            
            XSSFCell AdjResActcell = Actrow.getCell(5);
            String ActAdjRes=AdjResActcell.toString();
            
            String ActChkStr=ActClmNum + ":" + ActAdjRes;
            
               
            //Loop through Expected sheet
            for(int j=ExpFirstRow; j <= ExplastRow; j++) 
            {
                XSSFRow Exprow = ExpSheet.getRow(j);
                XSSFCell ClmNumExpcell = Exprow.getCell(0);
                String ExpClmNum=ClmNumExpcell.toString();
                
                XSSFCell AdjResExpcell = Exprow.getCell(5);
                String ExpAdjRes=AdjResExpcell.toString();
                
                String ExpChkStr=ExpClmNum + ":" + ExpAdjRes;
                
                if(ActChkStr.equalsIgnoreCase(ExpChkStr))
                {
                	rowexists=true;
                	break;
                }
            }
            
            if(!rowexists)
            {
            	  int firstExpcell = Actrow.getFirstCellNum();
                  int lastExpcell = Actrow.getLastCellNum();
                  
                  // Compare all cells in a row
                  aFinalData.append("<tr>");
                  aFinalData.append("<td style= \"background-color:red\">" + "Row Added in Actual:" + " </td>");
                  for(int l=firstExpcell; l <= lastExpcell-1; l++) 
                  {
                      XSSFCell Expcell = Actrow.getCell(l);
                      String cellval=Expcell.toString();
                      aFinalData.append("<td>" + cellval + "</td>");
                  }
                  aFinalData.append("</tr>");
                  equalSheets = false;
            }
        }
        return equalSheets;
    }
	
	
	public static boolean compareTwoSheets(XSSFSheet ExpSheet, XSSFSheet ActSheet, StringBuilder aDiffRows, StringBuilder ActMissingRows, StringBuilder ActAddedRows) throws IOException 
	{
        int ExpFirstRow = ExpSheet.getFirstRowNum(); 
        int ExplastRow = ExpSheet.getLastRowNum();
        boolean equalSheets = true;
        
        for(int i=ExpFirstRow; i <= ExplastRow; i++) 
        {
            XSSFRow Exprow = ExpSheet.getRow(i);
            XSSFRow Actrow = ActSheet.getRow(i);
            if(!compareTwoRows(Exprow, Actrow, aDiffRows)) 
            {
                equalSheets = false;
                System.out.println("Row "+i+" - Not Equal");
            } else 
            {
                System.out.println("Row "+i+" - Equal");
            }
        }
        return equalSheets;
    }

    public static boolean compareTwoRows(XSSFRow Exprow, XSSFRow Actrow,StringBuilder aFinalData) throws IOException 
    {
		
        if((Exprow == null) && (Actrow == null)) 
        {
            return true;
        } else if((Exprow == null) || (Actrow == null)) 
        {
            return false;
        }
        
        int firstExpcell = Exprow.getFirstCellNum();
        int lastExpcell = Exprow.getLastCellNum();
        boolean equalRows = true;
        

        StringBuilder aExpRowData = new StringBuilder();
      	StringBuilder aActRowData = new StringBuilder();

      	aExpRowData.append("<tr>");
      	aExpRowData.append("<td> Expected Row: </td>");
      	
      	aActRowData.append("<tr>");
      	aActRowData.append("<td> Actual Row: </td>");
        
        for(int i=firstExpcell; i <= lastExpcell-1; i++) 
        {
            XSSFCell Expcell = Exprow.getCell(i);
            XSSFCell Actcell = Actrow.getCell(i);
            String retvalue=compareTwoCells(Expcell, Actcell);
            
            String Expcellval=null;
            String Actcellval=null;
            
            if(Exprow.getCell(i) != null)
            {
            	 Expcellval=Expcell.toString();
            }
            if(Actrow.getCell(i) != null)
            {
            	Actcellval=Actcell.toString();
            }

            if(retvalue.contains("false")) 
            {
              	aExpRowData.append("<td style= \"background-color:green\">" + Expcellval + "</td>");
                aActRowData.append("<td style= \"background-color:red\">" + Actcellval + "</td>");
                equalRows = false;
            } else
            {
                aExpRowData.append("<td>" + Expcellval + "</td>");
                aActRowData.append("<td>" + Actcellval + "</td>");
            }
        }
        
      	aExpRowData.append("</tr>");
      	aActRowData.append("</tr>");
        
        if(!equalRows)
        {
        	aFinalData.append(aExpRowData.toString());
        	aFinalData.append(aActRowData.toString());
            aFinalData.append("<tr><td style= \"background-color:silver\" colspan=\"" + (lastExpcell+1) + "\""   + "</td></tr>");
        }
        
        return equalRows;
    }
    

    @SuppressWarnings({ "deprecation", "static-access" })
	public static String compareTwoCells(XSSFCell Expcell, XSSFCell Actcell) 
    {
    	String Expval = null;
    	String Actval = null;
    	
    	boolean equalcells=true;
    	
        if((Expcell == null || Expcell.getCellType() == Expcell.CELL_TYPE_BLANK) && (Actcell == null || Actcell.getCellType() == Actcell.CELL_TYPE_BLANK))
        {
            return "true";
        } 
        else if((Expcell == null) || (Actcell == null)) 
        {
            return "false";
        }
        
        int type1 = Expcell.getCellType();
        int type2 = Actcell.getCellType();
        if (type1 == type2)
        {

                switch (Expcell.getCellType()) 
                {
                case HSSFCell.CELL_TYPE_NUMERIC:
                    if (Expcell.getNumericCellValue() == Actcell.getNumericCellValue()) 
                    {
                    	double aExpval = Expcell.getNumericCellValue();
                    	double aActval = Actcell.getNumericCellValue();

                    	Expval=String.valueOf(aExpval);
                    	Actval=String.valueOf(aActval);
                    	
                    }
                    else
                    {
                    	double aExpval = Expcell.getNumericCellValue();
                    	double aActval = Actcell.getNumericCellValue();

                    	Expval=String.valueOf(aExpval);
                    	Actval=String.valueOf(aActval);
                    	
                    	equalcells=false;
                    }
                    break;
                case HSSFCell.CELL_TYPE_STRING:
                    if (Expcell.getStringCellValue().equals(Actcell.getStringCellValue())) 
                    {
                    	
                     	Expval = Expcell.getStringCellValue();
                    	Actval = Actcell.getStringCellValue();
                    	
                    }else
                    {
                    	Expval = Expcell.getStringCellValue();
                    	Actval = Actcell.getStringCellValue();
                    	
                    	equalcells=false;
                    }
                    break;
                case HSSFCell.CELL_TYPE_BLANK:
                    if (Actcell.getCellType() == HSSFCell.CELL_TYPE_BLANK) 
                    {
                    }
                    break;
                default:
                    if (Expcell.getStringCellValue().equals(Actcell.getStringCellValue())) 
                    {
                       	Expval = Expcell.getStringCellValue();
                    	Actval = Actcell.getStringCellValue();
                    }else
                    {
                    	Expval = Expcell.getStringCellValue();
                    	Actval = Actcell.getStringCellValue();
                    	
                    	equalcells=false;
                    }
                    break;
                }
            
        } else 
        {
            return Expval + "," + Actval + "," + "false";
            
        }
        
        if(equalcells)
        {
        	return Expval + "," + Actval + "," + "true";
        }else
        {
            return Expval + "," + Actval + "," + "false";
        }
        
    }
    
    
	public static void GetActualFile(String DBUsername, String DBPassword,String DBHostname, String DBServName,String ActFilePath, String aPlanCd,String aCycleDate) throws Exception 
    {
    	try
		  {
    			//Create Excel Workbook
    		   XSSFWorkbook aWorkbook = new XSSFWorkbook(); 
    	       XSSFSheet aSheet = aWorkbook.createSheet("ActualSheet");
    	       XSSFRow row=aSheet.createRow(0);
    	       XSSFCell cell;    		
		 
    	     //String aDBJdbcUrl = "jdbc:oracle:thin:aaa3205/VaApr2017123@florasit-scan:1521/FEPLSIT_RPTW";
    		  String aDBJdbcUrl = "jdbc:oracle:thin:" + DBUsername + "/" + DBPassword + "@" + DBHostname + ":1521/" + DBServName;
			  Connection conn = DriverManager.getConnection(aDBJdbcUrl);
			  if (conn != null) 
			  {
				 System.out.println("Connected to DB");
				// String aQuery="Select P.CLM_NBR_SHBR,P.DIRN_PAYMT_CD,P.EDIT_OVERR_CD,P.ENRL_CD, P.ADJDN_RESP_DSPOSN_CD,P.TRANS_ID_SHBR,CHK_OR_EFT_TRC_NBR, t.* from fepl_vchr.pav_gl_trans t, fepl_claim.clm_vchr_paymt p,(SELECT  VP.CLM_NBR_SHBR,VP.TRANS_ID_SHBR,VP.ADJDN_RESP_DSPOSN_CD,gl.pav_clm_cyc_trans_id,vp.clm_vers_id FROM  FEPL_claim.clm_vchr VC INNER JOIN  FEPL_CLAIM.CLM_VCHR_PAYMT VP on (VP.CLM_VCHR_ID =VC.CLM_VCHR_ID) INNER JOIN FEPL_VCHR.PAV_CLM_CYC_TRANS CT ON  (VP.CLM_NBR_SHBR = CT.DCN)and (VP.clm_vers_id = CT.CLM_Vers_ID) INNER join  FEPL_VCHR.PAV_GL_TRANS GL ON  (GL.PAV_CLM_CYC_TRANS_ID=CT.PAV_CLM_CYC_TRANS_ID) WHERE CT.PLAN_CD=? and to_char(GL.CYC_DT,'MM/DD/YYYY')=  ? )  g, FEPL_CLAIM.CLM_VCHR cc where g.pav_clm_cyc_trans_id=t.pav_clm_cyc_trans_id and g.trans_id_shbr = p.trans_id_shbr and g.clm_vers_id=p.clm_vers_id and CC.CLM_VCHR_ID = P.CLM_VCHR_ID order by p.CLM_NBR_SHBR,P.ADJDN_RESP_DSPOSN_CD asc ";
			
				// String aQuery="Select P.CLM_NBR_SHBR,P.DIRN_PAYMT_CD,P.EDIT_OVERR_CD,P.ENRL_CD, P.ADJDN_RESP_DSPOSN_CD,P.TRANS_ID_SHBR,CHK_OR_EFT_TRC_NBR, t.* from fepl_vchr.pav_gl_trans t, fepl_claim.clm_vchr_paymt p,(SELECT  VP.CLM_NBR_SHBR,VP.TRANS_ID_SHBR,VP.ADJDN_RESP_DSPOSN_CD,gl.pav_clm_cyc_trans_id,vp.clm_vers_id FROM  FEPL_claim.clm_vchr VC INNER JOIN  FEPL_CLAIM.CLM_VCHR_PAYMT VP on (VP.CLM_VCHR_ID =VC.CLM_VCHR_ID) INNER JOIN FEPL_VCHR.PAV_CLM_CYC_TRANS CT ON  (VP.CLM_NBR_SHBR = CT.DCN)and (VP.clm_vers_id = CT.CLM_Vers_ID) INNER join  FEPL_VCHR.PAV_GL_TRANS GL ON  (GL.PAV_CLM_CYC_TRANS_ID=CT.PAV_CLM_CYC_TRANS_ID) WHERE CT.PLAN_CD=? and to_char(GL.CYC_DT,'YYYYMMDD')=  ? )  g, FEPL_CLAIM.CLM_VCHR cc where g.pav_clm_cyc_trans_id=t.pav_clm_cyc_trans_id and g.trans_id_shbr = p.trans_id_shbr and g.clm_vers_id=p.clm_vers_id and CC.CLM_VCHR_ID = P.CLM_VCHR_ID order by p.CLM_NBR_SHBR,P.ADJDN_RESP_DSPOSN_CD asc ";
				   
				 String aQuery="Select P.CLM_NBR_SHBR,P.SUBR_ID, P.DIRN_PAYMT_CD, P.EDIT_OVERR_CD, P.ENRL_CD, P.ADJDN_RESP_DSPOSN_CD, T.ACCT_NBR,T.SUB_ACCT_NBR,T.ENT_CD,T.CJA_CD, T.RPT_CD,T.CORPT_DSPOSN_CD,T.GRP_NBR, T.GL_AMT,T.GL_DESC,T.GL_SORT_KEY from fepl_vchr.pav_gl_trans t, fepl_claim.clm_vchr_paymt p,(SELECT  VP.CLM_NBR_SHBR,VP.TRANS_ID_SHBR,VP.ADJDN_RESP_DSPOSN_CD,gl.pav_clm_cyc_trans_id,vp.clm_vers_id FROM  FEPL_claim.clm_vchr VC INNER JOIN  FEPL_CLAIM.CLM_VCHR_PAYMT VP on (VP.CLM_VCHR_ID =VC.CLM_VCHR_ID) INNER JOIN FEPL_VCHR.PAV_CLM_CYC_TRANS CT ON  (VP.CLM_NBR_SHBR = CT.DCN)and (VP.clm_vers_id = CT.CLM_Vers_ID) INNER join  FEPL_VCHR.PAV_GL_TRANS GL ON  (GL.PAV_CLM_CYC_TRANS_ID=CT.PAV_CLM_CYC_TRANS_ID) WHERE CT.PLAN_CD=? and to_char(GL.CYC_DT,'YYYYMMDD')=  ? )  g, FEPL_CLAIM.CLM_VCHR cc where g.pav_clm_cyc_trans_id=t.pav_clm_cyc_trans_id and g.trans_id_shbr = p.trans_id_shbr and g.clm_vers_id=p.clm_vers_id and CC.CLM_VCHR_ID = P.CLM_VCHR_ID order by p.CLM_NBR_SHBR,P.ADJDN_RESP_DSPOSN_CD asc ";
				 System.out.println("aQuery:" + aQuery);
				 
				 //String aQuery="Select P.CLM_NBR_SHBR,P.DIRN_PAYMT_CD,P.EDIT_OVERR_CD,P.ENRL_CD, P.ADJDN_RESP_DSPOSN_CD,P.TRANS_ID_SHBR,CHK_OR_EFT_TRC_NBR, t.* from fepl_vchr.pav_gl_trans t, fepl_claim.clm_vchr_paymt p,(SELECT  VP.CLM_NBR_SHBR,VP.TRANS_ID_SHBR,VP.ADJDN_RESP_DSPOSN_CD,gl.pav_clm_cyc_trans_id,vp.clm_vers_id FROM  FEPL_claim.clm_vchr VC INNER JOIN  FEPL_CLAIM.CLM_VCHR_PAYMT VP on (VP.CLM_VCHR_ID =VC.CLM_VCHR_ID) INNER JOIN FEPL_VCHR.PAV_CLM_CYC_TRANS CT ON  (VP.CLM_NBR_SHBR = CT.DCN)and (VP.clm_vers_id = CT.CLM_Vers_ID) INNER join  FEPL_VCHR.PAV_GL_TRANS GL ON  (GL.PAV_CLM_CYC_TRANS_ID=CT.PAV_CLM_CYC_TRANS_ID) WHERE CT.PLAN_CD=" + aPlanCd + "and to_char(GL.CYC_DT,'MM/DD/YYYY')="  + aCycleDate +" )  g, FEPL_CLAIM.CLM_VCHR cc where g.pav_clm_cyc_trans_id=t.pav_clm_cyc_trans_id and g.trans_id_shbr = p.trans_id_shbr and g.clm_vers_id=p.clm_vers_id and CC.CLM_VCHR_ID = P.CLM_VCHR_ID order by p.CLM_NBR_SHBR,P.ADJDN_RESP_DSPOSN_CD asc ";
				  PreparedStatement pstmt = conn.prepareStatement(aQuery);
				  pstmt.setString(1, aPlanCd);
				  pstmt.setString(2, aCycleDate); 
				  ResultSet rs = pstmt.executeQuery();
				  
				  //Get Column names and Set in Excel
				   ResultSetMetaData rsmd = rs.getMetaData();
				   int columnCount = rsmd.getColumnCount();
				   for (int i = 1; i <= columnCount; i++ ) 
				   {
				     String aColname = rsmd.getColumnName(i);
				     cell=row.createCell(i-1);
				   	 cell.setCellValue(aColname);
				   }
				   
				  //Set Data  in Excel
				  int aRowNum=1;
				  while (rs.next())
				  {
					  row=aSheet.createRow(aRowNum);
					  for (int i = 1; i <= columnCount; i++ ) 
					   {
						  cell=row.createCell(i-1);
						  cell.setCellValue(rs.getString(i));
					   }
					  aRowNum=aRowNum+1;
				  }
			  }
		
	       FileOutputStream out = new FileOutputStream(new File(ActFilePath));
	       aWorkbook.write(out);
	       out.close();
	       aWorkbook.close();
	       System.out.println("Actual WorkBook created from DB:{" +ActFilePath + "}"); 
		  }catch (Exception theException)
        	{
	          	theException.printStackTrace();
	          	throw theException;
	        }
    }

}
