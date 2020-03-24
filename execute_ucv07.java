import java.io.File;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelExec;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.Session;


public class execute_ucv07
{

	 public static void main(String[] args) throws Exception
	 {
       	 boolean aTestSuiteFailflag = false;

		/* String aExecEnv=args[0];
		 String theHostName=args[1];
		 String theUserName=args[2];
		 String thePassword=args[3];
		 String aCycleDate=args[4];
		 String aplancode=args[5];*/
		 

		//Read the Parameters
			File fXmlFile = new File("E:/Automation/pav/PAVRegression/InputParameters.xml");
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();
			NodeList nList = doc.getElementsByTagName("DownStreamValues");
			Node nNode=nList.item(0);
			Element eElement = (Element) nNode;
			String spCycleDate=eElement.getElementsByTagName("spCycleDate").item(0).getTextContent();
			String theHostName=eElement.getElementsByTagName("theHostName").item(0).getTextContent();
			String theUserName=eElement.getElementsByTagName("theUserName").item(0).getTextContent();
			String thePassword=eElement.getElementsByTagName("thePassword").item(0).getTextContent();
			String aExecEnv=eElement.getElementsByTagName("aExecEnv").item(0).getTextContent();
			String aplancode=eElement.getElementsByTagName("aplancode").item(0).getTextContent();
		
			
		 String theCommand="cd /opt/" + aExecEnv + "/bridge/script && " + "./pav_ucv07.sh " + aplancode + " " +  spCycleDate;
		 System.out.println("theCommand:" + theCommand);
		 final StringBuffer sb = new StringBuffer();
	 	 java.util.Properties config = new java.util.Properties();
	     config.put("StrictHostKeyChecking", "no");
	     JSch jsch = new JSch();
	     Session session=jsch.getSession(theUserName, theHostName, 22);
	     session.setPassword(thePassword);
	     session.setConfig(config);
	     session.connect();

	     Channel channel=session.openChannel("exec");
	     ((ChannelExec)channel).setCommand(theCommand);
	     channel.setInputStream(null);
	     ((ChannelExec)channel).setErrStream(System.err);

	     InputStream in=channel.getInputStream();
	     channel.connect();

	     byte[] tmp=new byte[1024];
	     while(true)
	     {
	       while(in.available()>0)
	       {
	         int i=in.read(tmp, 0, 1024);
	         if(i<0)break;
	         String aOutLog=new String(tmp, 0, i);
	         if(aOutLog.contains("Build Failed"))
	         {
	    			aTestSuiteFailflag=true;
	    			throw new JavaException("Build Failed When Running Test Suite...:" + aOutLog);
	         }
	         sb.append(aOutLog).append('\n');
	         System.out.println(new String(tmp, 0, i));

	         if(aTestSuiteFailflag)
	        	 break;
	       }

	       if(aTestSuiteFailflag)
	    	   break;

	       if(channel.isClosed())
	       {
	     	  System.out.println("exit-status: "+channel.getExitStatus());
	          break;
	       }
	       try{Thread.sleep(1000);}catch(Exception theException){System.out.println(theException);}
	     }

	     channel.disconnect();
	     session.disconnect();
	     System.out.println(theCommand  +  "Executed");

	  }
}


class JavaException extends Exception
{

	private static final long serialVersionUID = 1L;
	String str1;
	JavaException(String str2) {
       str1=str2;
    }

    public String toString(){
       return ("Output String = "+str1) ;
    }
}