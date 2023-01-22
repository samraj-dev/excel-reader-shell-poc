# Excel Reader

### Technologies
* Apache POI
* Spring Shell

### Build Jar
Run the following command to build the jar file
```
mvn clean install
```

### Read Excel 
Run the following command to start the shell
```
java.exe -jar excel-reader-shell-poc-0.0.1-SNAPSHOT.jar
```
A shell will be shown, run the below command to read and display the contents of the excel file
```read <filepath>``` example shown below
```read c:/excel/contact.xlsx```
