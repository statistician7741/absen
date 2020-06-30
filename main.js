const xlsx = require("node-xlsx").default;
const _ = require("lodash");
const fs = require("fs");
const moment = require('moment')
moment.locale('id')
const XlsxPopulate = require('xlsx-populate');
const file_path = __dirname + "/raw.xls";

const data = xlsx.parse(file_path);

let groups = {}

let date = moment();
const formatTanggalWithHour = 'HH:mm:ss M/D/YYYY'
const formatHour = 'HH:mm:ss'
const formatTglRawData = 'M/D/YYYY';

data[0].data.forEach((row, i, arr) => {
    if (i > 0) {
        if (!groups[row[2]]) {
            groups[row[2]] = {
                absen: {}
            };
        }
        date = moment(`${row[5]} ${row[4]}`, formatTanggalWithHour);
        
        const datang = moment(`${row[5]} ${row[4]}`, formatTanggalWithHour)
        const pulang = row[6] ? moment(`${row[8] ? row[8] : (row[7] ? row[7] : row[6])} ${row[4]}`, formatTanggalWithHour) : undefined
        //non ramadhan
        // let terlambat_menit = datang.diff(moment(datang).hour(7).minute(29).second(59), 'minutes');
        //ramadhan
        const ramadhan_start = moment("2019-05-05");
        const ramadhan_end = moment("2019-06-03");
        const isRamadhan = !(datang.isBefore(ramadhan_start) || datang.isAfter(ramadhan_end));
        let terlambat_menit = datang.diff(moment(datang).hour(7).minute(isRamadhan?59:29).second(59), 'minutes');

        let psw_menit = 0;

        if (pulang) {
            //non ramadhan
            // psw_menit = moment(pulang).hour(16).minute(pulang.day() === 5 ? 30 : 0).second(0).diff(pulang, 'minutes');
            //ramadhan
            psw_menit = moment(pulang).hour(isRamadhan?15:16).minute(pulang.day() === 5 ? 30 : 0).second(0).diff(pulang, 'minutes');
        } else{
            psw_menit = 999;
        }

        groups[row[2]].absen[row[4]] = {
            datang: {
                pukul: terlambat_menit < 510 ?datang.format(formatHour):'tidak absen',
                telat: terlambat_menit > 0 ? (terlambat_menit < 510?terlambat_menit:999) : '-'
            },
            pulang: {
                pukul: pulang ? pulang.format(formatHour) : (terlambat_menit < 509?'tidak absen':datang.format(formatHour)),
                kurang: psw_menit > 0 && terlambat_menit < 510 ? psw_menit : '-'
            },
        }

        if (terlambat_menit > 90) {
            groups[row[2]].absen[row[4]].TL4 = 'v'
        }
        if (terlambat_menit >= 1 && terlambat_menit <= 30) {
            groups[row[2]].absen[row[4]].TL1 = 'v'
        }
        if (terlambat_menit >= 31 && terlambat_menit <= 60) {
            groups[row[2]].absen[row[4]].TL2 = 'v'
        }
        if (terlambat_menit >= 61 && terlambat_menit <= 90) {
            groups[row[2]].absen[row[4]].TL3 = 'v'
        }
        if (psw_menit > 90) {
            groups[row[2]].absen[row[4]].PSW4 = terlambat_menit < 510?'v':'-'
        }
        if (psw_menit >= 1 && psw_menit <= 30) {
            groups[row[2]].absen[row[4]].PSW1 = 'v'
        }
        if (psw_menit >= 31 && psw_menit <= 60) {
            groups[row[2]].absen[row[4]].PSW2 = 'v'
        }
        if (psw_menit >= 61 && psw_menit <= 90) {
            groups[row[2]].absen[row[4]].PSW3 = 'v'
        }
    }
})

XlsxPopulate.fromFileAsync(__dirname + "/rekap.xlsx")
    .then(workbook => {
        let index = 0;
        for (let nama in groups) {
            if (groups.hasOwnProperty(nama)) {
                if (true) {
                    let sheet = workbook.sheet(index);
                    sheet.name(nama)
                    workbook.sheet(index).cell("B1").value(date.format('MMMM YYYY'));
                    workbook.sheet(index).cell("B2").value(nama);
                    let row = 7
                    for (let i = 1; i <= date.endOf('month').date(); i++) {
                        let r = sheet.range('A' + row + ':N' + row);
                        (moment(date).date(i).day() === 0 || moment(date).date(i).day() === 6)&&r.style("fill", {
                            type: "solid",
                            color: {
                                rgb: "8c8c8c"
                            }
                        })
                        let data = groups[nama].absen[moment(date).date(i).format(formatTglRawData)];
                        let arr = [
                            moment(date).date(i).format('DD/MM/YYYY'),
                            moment(date).date(i).format('dddd'),
                            data ? data.datang.pukul : '-',
                            data ? data.datang.telat : '-',
                            data ? data.pulang.pukul : '-',
                            data ? data.pulang.kurang : '-',

                            data ? (data.TL1?data.TL1:'-') : '-',
                            data ? (data.TL2?data.TL2:'-') : '-',
                            data ? (data.TL3?data.TL3:'-') : '-',
                            data ? (data.TL4?data.TL4:'-') : '-',

                            data ? (data.PSW1?data.PSW1:'-') : '-',
                            data ? (data.PSW2?data.PSW2:'-') : '-',
                            data ? (data.PSW3?data.PSW3:'-') : '-',
                            data ? (data.PSW4?data.PSW4:'-') : '-',
                        ]
                        r.value([
                            arr
                        ]);
                        row++
                    }
                }
                index++;
            }
        }

        if (fs.existsSync(__dirname + `/rekap_ok.xlsx`)) {
            fs.unlinkSync(__dirname + `/rekap_ok.xlsx`);
        }
        workbook.toFileAsync(__dirname + `/rekap_ok.xlsx`);
    }).then(dataa => {
        console.log('Finished');
    })