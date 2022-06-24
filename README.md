# PDFReplicator


This Application is built to help people generate millions of PDFs from a HTML template by populating data from 
XLS or XLSX file. This was preprared with the purpose of generating multiple copies of a letter/document that needs 
to be peronalised for the receiver with their name, contact details and details associated with them. A typical example 
would be generating letters with common intent for individual customers personalized for each of them such as information 
about a new product or a new service or intent to extend the service terms and period.

## Requirements:
* Java 8 Runtime (minimum)

## At present a release is not uploaded. So, the users shall build the application at their end and use the application.
## I would be uploading the 0.9v of the applciation very soon.
## Steps to build and use the Application:
* Follow the instaructions from [https://gradle.org/install/](https://gradle.org/install/)
  to install Gradle in your machine.
* Run the command `./gradlew build` to build the applicaiton
* The packaged .JAR file is available in `./build/libs/` directory
* A script to run the built .JAR files are provided in the project root folder.
  * `PDFReplicator_v1_0.sh` for MAC and Linux machines. They need to be made as executable files before running them using the terminal command `./PDFReplicator_v1_0.sh`.
  * `PDFReplicator_v1_0.bat` for Windows PC.
* Alternately they can be run directly from terminal using the `java -Xms4g -Xmx5g -jar ./build/libs/app-1.0-all.jar` command. The `Xms` is to inform the OS to allocate 4gb RAM with a maximum range set as 5gb RAM using the tag `Xmx`. This was set to enable generating millions of PDFs using this application. 

This Application uses iText 7 core library along with pdfHTML plugin to convert HTML to PDF.
The iText 7 libraries are used under open source GNU AGPL license and I have made the source code of the application 
open source honouring the GNU AGPL license terms of the iText 7 library. The libraries are included in the build 
using the Maven repository through the Gradle build. You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see [License](https://itextpdf.com/en/how-buy/legal/agpl-gnu-affero-general-public-license.

