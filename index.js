const xl = require('excel4node');

exports.kpiExport = (req, res) => {
    res.set('Access-Control-Allow-Origin', '*');
    // res.set('Access-Control-Allow-Headers', 'Content-Type');
    // res.set('Access-Control-Allow-Methods', 'POST');

    if (req.method === 'POST') {
        let headers = req.body.headers;
        let data = req.body.data;
        let channelList = req.body.channelList;
        let ws = [];
        let wb = new xl.Workbook();
        let myStyle = wb.createStyle({
            font: {
                bold: true,
                size: 14
            },
            alignment: {
                wrapText: true,
                horizontal: 'center',
                vertical: 'center'
            },
            border: {
                left: {
                    style: 'thin',
                    color: 'gray-50'
                },
                right: {
                    style: 'thin',
                    color: 'gray-50'
                },
                top: {
                    style: 'thin',
                    color: 'gray-50'
                },
                bottom: {
                    style: 'thin',
                    color: 'gray-50'
                },
                diagonal: {
                    style: 'thin',
                    color: 'gray-50'
                },
            },
            fill: {
                type: 'pattern',
                patternType: 'solid',
                bgColor: 'light yellow',
                fgColor: 'light yellow'
            }
        });

        for (let n = 0; n < channelList.length; n++) {
            ws[n] = wb.addWorksheet(channelList[n].name);

            for (let i = 1; i <= headers[n].length; i++) {
                ws[n].cell(1, i)
                    .string(headers[n][i - 1].text)
                    .style(myStyle);

                for (let j = 1; j <= data[n].length; j++) {
                    for (let k = 1; k <= headers[n].length; k++) {
                        let X = '';
                        X = String(data[n][j - 1][headers[n][k - 1].value]);
                        ws[n].cell(j + 1, k).string(X).style({
                            alignment: {
                                horizontal: 'right'
                            }
                        });
                    }
                }
            }
        }
        wb.write('kpi_table.xlsx', res);
    } else {
        // Set CORS headers for the main request
        res.set('Access-Control-Allow-Origin', '*');
    }
};
