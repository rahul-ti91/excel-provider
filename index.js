var excel = require('node-excel-export'),
    express = require('express'),
    cfenv = require("cfenv"),
    request = require('request');

app = express();
var appEnv = cfenv.getAppEnv();

var PORT = 3001;
(function () {
    if (appEnv.isLocal) {

    } else {
        PORT = process.env.PORT;
    }
    console.log(PORT);
})();

var styles = {
    headerDark: {
        fill: {
            fgColor: {
                rgb: 'FF455A64;'
            }
        },
        font: {
            color: {
                rgb: 'FFFFFFFF'
            },
            sz: 13,
            bold: false,
            underline: false
        },
        alignment: {
            horizontal: "center"
        }
    },
    cellPink: {
        alignment: {
            horizontal: "center"
        }
    },
    cellDate: {
        alignment: {
            horizontal: "center"
        },
        numFmt: "m/dd/yy"
    },
    cellApp: {
        fill: {
            fgColor: {
                rgb: 'FF00695C'
            }
        },
        font: {
            color: {
                rgb: 'FFFFFFFF'
            },
            sz: 13,
            bold: false,
            italic: true,
            underline: false
        },
        alignment: {
            horizontal: "center"
        }
    }
};

var appObj = { // <- the key should match the actual data key 
    displayName: 'Application', // <- Here you specify the column header 
    headerStyle: styles.headerDark, // <- Header style 
    cellStyle: styles.cellApp,
    width: 120 // <- width in pixels 
};

var normalObj = function (displayName, size, centerAlign) {
    if (centerAlign) {
        return {
            displayName: displayName,
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink,
            width: size
        }
    } else {
        return {
            displayName: displayName,
            headerStyle: styles.headerDark,
            width: size
        }
    }
};

var dateObj = function (displayName, size) {
    return {
        displayName: displayName,
        headerStyle: styles.headerDark,
        cellStyle: styles.cellDate,
        cellFormat: function (value, row) {
            return (new Date(value));
        },
        width: size
    }
}

var specification_colseactivity = {
    applicationName: appObj,
    activityDate: dateObj("Closing Date", 110),
    frequency: normalObj("Frequency", 110, true),
    description: normalObj("Description", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

var specification_outages = {
    applicationName: appObj,
    outageDate: dateObj("Outage Date", 110),
    startTime: normalObj("Start Time", 110, true),
    duration: normalObj("Duration", 110, true),
    outageType: normalObj("Outage Type", 110, true),
    rcaDone: normalObj("RCA Status", 110, true),
    outageReason: normalObj("Outage Reason", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

var specification_releasecalendar = {
    applicationName: appObj,
    releaseCompletionDate: dateObj("Release Completion Date", 190),
    upcomingReleaseDate: dateObj("Upcoming Release Date", 190),
    creationDate: dateObj("Creation Date", 110)
};

var specification_drcalendar = {
    applicationName: appObj,
    drCompletionDate: dateObj("DR Completion Date", 150),
    upcomingDRDate: dateObj("Upcoming DR Date", 150),
    creationDate: dateObj("Creation Date", 110)
};

var specification_appreciations = {
    applicationName: appObj,
    appreciation: normalObj("Appreciations", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

var specification_coresissues = {
    applicationName: appObj,
    issue: normalObj("Issue", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

var specification_ideas = {
    applicationName: appObj,
    ideaState: normalObj("Idea State", 110, true),
    ideaDescription: normalObj("Idea Description", 220, false),
    businessBenefits: normalObj("Business Benefits", 220, false),
    implamentationPlan: normalObj("Implamentation Plan", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

var specification_trainings = {
    empName: normalObj("Employee Name", 110, true),
    trainingType: normalObj("Training Type", 110, true),
    trainingName: normalObj("Training Name", 110, true),
    creationDate: dateObj("Creation Date", 110)
};

var specification_nonsndata = {
    applicationName: appObj,
    week: dateObj("Week", 110),
    data: normalObj("Data", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

var specification_weeklyhighlights = {
    applicationName: appObj,
    week: dateObj("Week", 110),
    highlights: normalObj("Highlights", 220, false),
    creationDate: dateObj("Creation Date", 110)
};

app.get('/Large', function (req, res) {

    request('https://cores-msr-jpa.run.aws-usw02-pr.ice.predix.io/report?month=' + (req.query.month | 3), function (error, response, body) {
        console.log('error:', error); // Print the error if one occurred 
        console.log('statusCode:', response && response.statusCode); // Print the response status code if a response was received 
        console.log('body:', body);
        var result = JSON.parse(body);
        var report = excel.buildExport(
            [
                {
                    name: 'Closing Activity', // <- Specify sheet name (optional)  
                    specification: specification_colseactivity, // <- Report specification 
                    data: result.closeActivities // <-- Report data 
                }, {
                    name: 'Appreciations', // <- Specify sheet name (optional) 
                    specification: specification_appreciations, // <- Report specification 
                    data: result.appreciations // <-- Report data 
                }, {
                    name: 'Issues', // <- Specify sheet name (optional)  
                    specification: specification_coresissues, // <- Report specification 
                    data: result.coresIssues // <-- Report data 
                }, {
                    name: 'DR Calendar', // <- Specify sheet name (optional) 
                    specification: specification_drcalendar, // <- Report specification 
                    data: result.drcalendars // <-- Report data 
                }, {
                    name: 'Release Calendar', // <- Specify sheet name (optional)  
                    specification: specification_releasecalendar, // <- Report specification 
                    data: result.releaseCalendars // <-- Report data 
                }, {
                    name: 'Ideas', // <- Specify sheet name (optional) 
                    specification: specification_ideas, // <- Report specification 
                    data: result.ideas // <-- Report data 
                }, {
                    name: 'Non SN Data', // <- Specify sheet name (optional) 
                    specification: specification_nonsndata, // <- Report specification 
                    data: result.nonsnDatas // <-- Report data 
                }, {
                    name: 'Outages', // <- Specify sheet name (optional) 
                    specification: specification_outages, // <- Report specification 
                    data: result.outages // <-- Report data 
                }, {
                    name: 'Trainings', // <- Specify sheet name (optional) 
                    specification: specification_trainings, // <- Report specification 
                    data: result.trainings // <-- Report data 
                }
            ]);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats');
        res.setHeader("Content-Disposition", "attachment; filename=" + "report.xlsx");
        res.attachment('report.xlsx'); // This is sails.js specific (in general you need to set headers) 
        res.send(report);
    });
});
app.listen(PORT);
console.log('Listening on port ' + PORT);
