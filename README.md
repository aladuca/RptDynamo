RptDynamo
=========

Crystal Reports Generator

Command line options
- -c configuration file path
- -j job file path
- -v verbose console output

Example Config File
-------------------
```
{
	"smtp" : {
		"server": "smtp.example.com",
		"port": "587",
		"username": "smtpuser",
		"password": "smtppassword",
		"sender": "donotreply@smtp.example.com",
		"ssl": "true"
		},
	"queue": {
		"type": "msmq",
		"name": "queue name"
		}
}
		
```

Example Report Job File
-----------------------
Report with multiple parameters.
```
{
  "report": {
    "Filename": "full path to file",
    "parameter": [
      {
        "Name": "1st parameter name",
        "text": "value"
      },
      {
        "Name": "2nd parameter name",
        "text": "value"
      },
      {
        "Name": "3rd parameter name",
        "text": "value"
      }
    ],
    "output": "pdf"
  },
  "email" : {
	"to": ["usera1@example.com", "usera2@example.com"],
	"cc": ["userb1@example.com", "userb2@example.com"]
	}
}
```
Report with no parameters
```
{
  "report": {
    "Filename": "full path to file",
    "output": "pdf"
  },
  "email" : {
	"to": ["usera1@example.com", "usera2@example.com"],
	"cc": ["userb1@example.com", "userb2@example.com"]
	}
}
```
