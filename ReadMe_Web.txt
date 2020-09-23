Readme_WEB File.
-------------
Created by Peter Verburgh.

Remark  : i use a freeware ocx for uploading files FTP .. EZFTP.OCX
          I've posted the EZFTP.OCX but it wasn't included in the zip file...
		  
		  You can also get the OCX file on my homepage.. (also under construction)
		  http://users.skynet.be/verburgh.peter

Remark 2 : If you have some source code for how i can detect how multiple IP are installed on
           your computer.. then i could change the code.. 
		   because , winsock takes the default network card IP..
		   and if you have a modem & networkcard.. it will detects the default card..
			email : peter.verburgh2@yucom.be

Remark 3 : If you find bugs , please msg me !


Important Remark : Please Vote For me ..

How it works..
---------------
If you start this program and fill up the settings.. url FTP, local file..,
remote file.. login & pasw for FTP upload..
And by Clicking on Start , my program detects the the IP address, opens the local file 
and replace the IPADR variable.. example to 217.134.11.11 
When its done , its uploads the file to your free hosting server..
So , if some-one's goes to your free - site , he/she will be redirected to your 
personal webserver.
Look at the code i have included in the Zip file  ... index.html  (javascript code)
If you go Offline , you have to click on stop , then the IP in the local file will
be changed to 'NONE'  and it sends that new file (index.html) to the free server.
If some-one's visit your site , he get a msgbox "Server is not online"
and will be directed to index2.html.. (you can choose...)



Install the program..
Once it's installed, you have to change the standard file on the web
server ISP (index.html)
Rename the index.html file to something else.. (index2.html).
Once it's done , copy the index.html file that from this package.
to your ISP webpage.

the content of the index.html file from this package is:



<SCRIPT LANGUAGE="javascript">

var IPADR ='NONE';

if(IPADR =='NONE')
{
// if your OWN webserver isn't active.. 
alert("Server is not active");
// Go to your other html file
// example....
// navigate("/index2.html");
} 
else
{
//Go to your OWN Webserver..
navigate("http://"+IPADR);
// : if you server not runs on port 80 .. example ..
// navigate("http://"+IPADR+":5080");   
// if your server runs on port 5080...
}

</SCRIPT>

Greetings.

Peter Verburgh.