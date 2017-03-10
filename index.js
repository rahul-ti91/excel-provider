var excel = require('node-excel-export'),
    express = require('express'),
    cfenv = require("cfenv"),
    request = require('request');
// You can define styles as json object 
// More info: https://github.com/protobi/js-xlsx#cell-styles 
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

app.get('/Large', function (req, res) {

    // The data set should have the following shape (Array of Objects) 
    // The order of the keys is irrelevant, it is also irrelevant if the 
    // dataset contains more fields as the report is build based on the 
    // specification provided above. But you should have all the fields 
    // that are listed in the report specification 
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
    //Array of objects representing heading rows (very top) 
    var heading = [
  [{
            value: 'a1',
            style: styles.headerDark
        }, {
            value: 'b1',
            style: styles.headerDark
        }, {
            value: 'c1',
            style: styles.headerDark
        }]
  , ['a2', 'b2', 'c2'] // <-- It can be only values 
];

    //Here you specify the export structure 
    var specification_colseactivity = {

        applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        activityDate: {
            displayName: 'Closing Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 110 // <- width in pixels 
        },
        frequency: {
            displayName: 'Frequency',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        description: {
            displayName: 'Description',
            headerStyle: styles.headerDark,
          //  cellStyle: styles.cellPink, // <- Cell style 
            width: 220 // <- width in pixels 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            }, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    var specification_outages = {
        applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        outageDate: {
            displayName: 'Outage Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        startTime: {
            displayName: 'Start Time',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        },
        duration: {
            displayName: 'Duration',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        },
        outageType: {
            displayName: 'Outage Type',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        },
        rcaDone: {
            displayName: 'RCA Status',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        },
        outageReason: {
            displayName: 'Outage Reason',
            headerStyle: styles.headerDark,
            //cellStyle: styles.cellPink, // <- Cell style 
            width: 220 // <- width in pixels 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 110 // <- width in pixels 
        }
    };

    var specification_releasecalendar = {
        applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        releaseCompletionDate: {
            displayName: 'Release Completion Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 190 // <- width in chars (when the number is passed as string) 
        },
        upcomingReleaseDate: {
            displayName: 'Upcoming Release Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 190 // <- width in pixels 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 110 // <- width in pixels 
        }
    };

    var specification_drcalendar = {
        applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        drCompletionDate: {
            displayName: 'DR Completion Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 150 // <- width in chars (when the number is passed as string) 
        },
        upcomingDRDate: {
            displayName: 'Upcoming DR Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 150 // <- width in pixels 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 110 // <- width in pixels 
        }
    };

    var specification_appreciations = {
         applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        appreciation: {
            displayName: 'Appreciations',
            headerStyle: styles.headerDark,
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    var specification_coresissues = {
         applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        issue: {
            displayName: 'Issue',
            headerStyle: styles.headerDark,
            width: 220 // <- width in chars (when the number is passed as string) 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    var specification_ideas = {
         applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        ideaState: {
            displayName: 'Idea State',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        ideaDescription: {
            displayName: 'Idea Description',
            headerStyle: styles.headerDark,
            width: 220 // <- width in chars (when the number is passed as string) 
        },
        businessBenefits: {
            displayName: 'Business Benefits',
            headerStyle: styles.headerDark,
            width: 220 // <- width in chars (when the number is passed as string) 
        },
        implamentationPlan: {
            displayName: 'Implamentation Plan',
            headerStyle: styles.headerDark,
            width: 220 // <- width in chars (when the number is passed as string) 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    var specification_trainings = {
        empName: {
            displayName: 'Employee Name',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        trainingType: {
            displayName: 'Training Type',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        trainingName: {
            displayName: 'Training Name',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    var specification_nonsndata = {
        applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        week: {
            displayName: 'Week',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        data: {
            displayName: 'Data',
            headerStyle: styles.headerDark,
            width: 220 // <- width in chars (when the number is passed as string) 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    var specification_weeklyhighlights = {
         applicationName: { // <- the key should match the actual data key 
            displayName: 'Application', // <- Here you specify the column header 
            headerStyle: styles.headerDark, // <- Header style 
            cellStyle:styles.cellApp,
            width: 120 // <- width in pixels 
        },
        week: {
            displayName: 'Week',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style 
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            width: 110 // <- width in chars (when the number is passed as string) 
        },
        highlights: {
            displayName: 'Highlights',
            headerStyle: styles.headerDark,
            width: 220 // <- width in chars (when the number is passed as string) 
        },
        creationDate: {
            displayName: 'Creation Date',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property 
                return (new Date(value)).toLocaleDateString();
            },
            cellStyle: styles.cellPink, // <- Cell style 
            width: 110 // <- width in pixels 
        }
    };

    request('https://cores-msr-jpa.run.aws-usw02-pr.ice.predix.io/report?month=' + (req.query.month | 3), function (error, response, body) {
        console.log('error:', error); // Print the error if one occurred 
        console.log('statusCode:', response && response.statusCode); // Print the response status code if a response was received 
        console.log('body:', body);
        var result = JSON.parse(body);
        var report = excel.buildExport(
            [
                {
                    name: 'Closing Activity', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_colseactivity, // <- Report specification 
                    data: result.closeActivities // <-- Report data 
                }, {
                    name: 'Appreciations', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_appreciations, // <- Report specification 
                    data: result.appreciations // <-- Report data 
                }, {
                    name: 'Issues', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_coresissues, // <- Report specification 
                    data: result.coresIssues // <-- Report data 
                }, {
                    name: 'DR Calendar', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_drcalendar, // <- Report specification 
                    data: result.drcalendars // <-- Report data 
                }, {
                    name: 'Release Calendar', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_releasecalendar, // <- Report specification 
                    data: result.releaseCalendars // <-- Report data 
                }, {
                    name: 'Ideas', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_ideas, // <- Report specification 
                    data: result.ideas // <-- Report data 
                }, {
                    name: 'Non SN Data', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_nonsndata, // <- Report specification 
                    data: result.nonsnDatas // <-- Report data 
                }, {
                    name: 'Outages', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
                    specification: specification_outages, // <- Report specification 
                    data: result.outages // <-- Report data 
                }, {
                    name: 'Trainings', // <- Specify sheet name (optional) 
                    //  heading: heading, // <- Raw heading array (optional) 
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
