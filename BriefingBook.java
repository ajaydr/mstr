package com.coke;
import com.microstrategy.web.objects.WebIServerSession;
import com.microstrategy.webapi.EnumDSSXMLAuthModes;
import com.microstrategy.web.objects.WebObjectsException;
import com.microstrategy.web.objects.WebObjectsFactory;
import com.microstrategy.web.objects.WebPrompts;
import com.microstrategy.web.objects.rw.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ResourceBundle;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;

public class BriefingBook {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	
		String dbURL;
		String mstrWebAppFolderLoc;
		String pdfFileFolder;
		String projectName;
		String userName;
		String pwd;
		String iServerName;
		String webRootURL;
		
		String pdfFileName ="Doc";
		RWInstance rwi=null;
		File file;
		FileOutputStream fos=null;
		WebIServerSession serverSession=null;
		WebObjectsFactory objectFactory=null;

		//String promptAnsXml1="<rsl><pa pt='3' pin='0' did='2294C41E44244ADAED930CB4E00A9DA0' tp='10'>Northeast</pa></rsl>";
		//String promptAnsXml2="<rsl><pa pt='3' pin='0' did='2294C41E44244ADAED930CB4E00A9DA0' tp='10'>Web</pa></rsl>";
		String getBookSQL="select TOP 1 BookID, BookName, BookStatus,BookUserID from dbo.BriefingBookRequest where BookStatus='Submitted' order by BookID ASC";
		String getBookDocsSQL="select DocId, DocParamStr from dbo.BriefingBookReqDocs where BookID=";
		String updateBookSQL="update dbo.BriefingBookRequest set BookStatus='In-Progress' where BookID=";
		String updateBookRevertSQL="update dbo.BriefingBookRequest set BookStatus='Submitted' where BookID=";
		String updateBookCompletionSQL="update dbo.BriefingBookRequest set BookStatus='Complete', BookFileURL='<a href=''#WEBROOTURL#/#PDFFILEFOLDER#/#FILE_NAME#'' target=''_blank''>Click</a>' where BookID=";

		int numPdfs=0;
		String BookID=null;
		String BookName=null;
		Connection conn = null;
		Statement stmt=null;
		int i=0;
		try {
			ResourceBundle propertiesFile = ResourceBundle.getBundle("BriefingBook");
			dbURL=propertiesFile.getString("dbURL");
			mstrWebAppFolderLoc=propertiesFile.getString("mstrWebAppFolderLoc")+"/";
			pdfFileFolder=propertiesFile.getString("pdfFileFolder")+"/";
			projectName=propertiesFile.getString("projectName");
			userName=propertiesFile.getString("userName");
			pwd=propertiesFile.getString("pwd");
			iServerName=propertiesFile.getString("iServerName");
			webRootURL=propertiesFile.getString("webRootURL");

			conn = DriverManager.getConnection(dbURL);
			stmt = conn.createStatement();
            ResultSet rs;
            rs = stmt.executeQuery(getBookSQL);
            while ( rs.next() ) {
            	BookID = rs.getString("BookID");
            	BookName = rs.getString("BookName");
                System.out.println("Processing Book Name: ID "+BookName+" : "+BookID);
            }
            if(BookID != null)
            {
            	stmt.executeUpdate(updateBookSQL+BookID);
            
            	objectFactory = WebObjectsFactory.getInstance();
            	serverSession = objectFactory.getIServerSession();
            	//serverSession.setApplicationType(EnumDSSXMLApplicationType.DssXmlApplicationCustomApp);
            	serverSession.setAuthMode(EnumDSSXMLAuthModes.DssXmlAuthStandard);
            	serverSession.setProjectName(projectName);
            	serverSession.setLogin(userName);
            	serverSession.setPassword(pwd);
            	serverSession.setServerName(iServerName);

            	rs=stmt.executeQuery(getBookDocsSQL+BookID);
            	i=0;
            	while ( rs.next() ) {
            		i++;
            		String DocID = rs.getString("DocID");
            		String promptAnsXml=rs.getString("DocParamStr");
            		System.out.println("Generating Document:"+DocID+"\n with promptxml:"+promptAnsXml);
  
            		if(DocID != null){
	            		rwi = objectFactory.getRWSource().getNewInstance(DocID);
	        			//rwi.setAsync(false);
	        			rwi.setMaxWait(-1);
	        			rwi.setPollingFrequency(250); 
	        			int status = rwi.pollStatus(); 
	        			rwi.setAsync(false);
	        			rwi.setExecutionMode(EnumRWExecutionModes.RW_MODE_PDF);
	        			//Ans prompts
	        			WebPrompts prompts = rwi.getPrompts();
	        			prompts.populateAnswers(promptAnsXml);
	        			prompts.validate();
	        			prompts.answerPrompts();
	        			//System.out.println("Prompts answered");
	        			//RWExportSettings expSet = rwi.getExportSettings();
	        			//expSet.setGridKey("NODEKEY_OF_THE_GRID");
	        			file = new File(mstrWebAppFolderLoc+pdfFileFolder+pdfFileName+i+".pdf");
	        			// if file doesn't exists, then create it
	        			if (!file.exists()) {
	        				file.createNewFile();
	        			}
	        			fos = new FileOutputStream(file);
	        			rwi.setMaxWait(-1);
	        			rwi.setPollingFrequency(250); 
	        			status = rwi.pollStatus();
	        			fos.write(rwi.getExportData());
	        			numPdfs++; 
	        			fos.flush();
	        			fos.close();
            		}
            		
            	}
           
			serverSession.closeSession();
			PDFMergerUtility PDFmerger = new PDFMergerUtility();
			PDFmerger.setDestinationFileName(mstrWebAppFolderLoc+pdfFileFolder+BookName+".pdf");
			PDDocument pfddoc;
			i=0;
			System.out.println("Number of individual Pdfs to be merged:"+numPdfs);
			for(int j=0 ; j<numPdfs; j++)
			{
				i++;
				file = new File(mstrWebAppFolderLoc+pdfFileFolder+pdfFileName+i+".pdf");
				pfddoc = PDDocument.load(file);
				PDFmerger.addSource(file);
				pfddoc.close();
			}
			PDFmerger.mergeDocuments(null);
			System.out.println("Documents merged complete");
		
			updateBookCompletionSQL =updateBookCompletionSQL.replace("#WEBROOTURL#", webRootURL).replace("#PDFFILEFOLDER#", pdfFileFolder).replace("#FILE_NAME#", URLEncoder.encode(BookName,"UTF-8")+".pdf")+BookID;
		    stmt.executeUpdate(updateBookCompletionSQL);
          }
		      
		} catch (IOException  | WebObjectsException | SQLException e) {
			if(stmt != null && BookID != null)
				try {
					stmt.executeUpdate(updateBookRevertSQL+BookID);
				} catch (SQLException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			e.printStackTrace();
		}
		finally{
			
			if(fos != null)
				try {
					fos.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			
			if(serverSession != null)
				try {
					serverSession.closeSession();
				} catch (WebObjectsException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			try {
                if (conn != null && !conn.isClosed()) {
                    conn.close();
                }
            } catch (SQLException ex) {
                ex.printStackTrace();
            }
		      
		     
		}
	}

}
