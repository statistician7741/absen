const xlsx = require("node-xlsx").default;
const _ = require("lodash");
const fs = require("fs");
const moment = require('moment')
moment.locale('id')
const XlsxPopulate = require('xlsx-populate');
const file_path = __dirname + "/raw.xls";
const shift_ppnpn = require('./config/env.config').shift_ppnpn

const data = xlsx.parse(file_path);

let groups = {}

let current_day = moment();
const today_id = moment().format('YYYY_MM_DD')
const yest_id = moment().subtract(1, 'day').format('YYYY_MM_DD')
const formatTanggalWithHour = 'HH:mm:ss M/D/YYYY'
const formatHour = 'HH:mm:ss'
const formatTglRawData = 'M/D/YYYY';

let ppnpns = {}
data[0].data.forEach((row, i, arr) => {
    if (i > 0) {
        if (!groups[row[2]]) {
            groups[row[2]] = {
                absen: {}
            };
        }
        current_day = moment(`${row[5]} ${row[4]}`, formatTanggalWithHour);
        for (let index = 5; index <= 8; index++) {
            if (row[index]) {
                if (!ppnpns[row[2]]) ppnpns[row[2]] = {}
                if (!ppnpns[row[2]][row[4]]) ppnpns[row[2]][row[4]] = {}
                if (!ppnpns[row[2]][row[4]].presensiArray) ppnpns[row[2]][row[4]].presensiArray = []
                ppnpns[row[2]][row[4]].presensiArray.push(moment(`${row[index]} ${row[4]}`, formatTanggalWithHour))
            }
        }

        const datang = moment(`${row[5]} ${row[4]}`, formatTanggalWithHour)
        const mid = moment(`${row[5]} ${row[4]}`, formatTanggalWithHour)
        const pulang = row[6] ? moment(`${row[8] ? row[8] : (row[7] ? row[7] : row[6])} ${row[4]}`, formatTanggalWithHour) : undefined
        //non ramadhan
        // let terlambat_menit = datang.diff(moment(datang).hour(7).minute(29).second(59), 'minutes');
        //ramadhan
        const ramadhan_start = moment("2019-05-05");
        const ramadhan_end = moment("2019-06-03");
        const isRamadhan = !(datang.isBefore(ramadhan_start) || datang.isAfter(ramadhan_end));
        let terlambat_menit = datang.diff(moment(datang).hour(7).minute(isRamadhan ? 59 : 29).second(59), 'minutes');

        let psw_menit = 0;

        if (pulang) {
            //non ramadhan
            // psw_menit = moment(pulang).hour(16).minute(pulang.day() === 5 ? 30 : 0).second(0).diff(pulang, 'minutes');
            //ramadhan
            psw_menit = moment(pulang).hour(isRamadhan ? 15 : 16).minute(pulang.day() === 5 ? 30 : 0).second(0).diff(pulang, 'minutes');
        } else {
            psw_menit = 999;
        }

        groups[row[2]].absen[row[4]] = {
            datang: {
                pukul: terlambat_menit < 510 ? datang.format(formatHour) : 'tidak absen',
                telat: terlambat_menit > 0 ? (terlambat_menit < 510 ? terlambat_menit : 999) : '-'
            },
            pulang: {
                pukul: pulang ? pulang.format(formatHour) : (terlambat_menit < 509 ? 'tidak absen' : datang.format(formatHour)),
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
            groups[row[2]].absen[row[4]].PSW4 = terlambat_menit < 510 ? 'v' : '-'
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

console.log(ppnpns);

XlsxPopulate.fromFileAsync(__dirname + "/rekap_ppnpns.xlsx")
    .then(workbook => {
        let index = 0;
        for (let nama in groups) {
            if (groups.hasOwnProperty(nama)) {
                if (true) {
                    let sheet = workbook.sheet(index);
                    sheet.name(nama)
                    workbook.sheet(index).cell("B1").value(current_day.format('MMMM YYYY'));
                    workbook.sheet(index).cell("B2").value(nama);
                    let row = 7
                    for (let i = 1; i <= current_day.endOf('month').date(); i++) {
                        let r = sheet.range('A' + row + ':N' + row);
                        (moment(current_day).date(i).day() === 0 || moment(current_day).date(i).day() === 6) && r.style("fill", {
                            type: "solid",
                            color: {
                                rgb: "8c8c8c"
                            }
                        })
                        let data = groups[nama].absen[moment(current_day).date(i).format(formatTglRawData)];
                        let arr = [
                            moment(current_day).date(i).format('DD/MM/YYYY'),
                            moment(current_day).date(i).format('dddd'),
                            data ? data.datang.pukul : '-',
                            data ? data.datang.telat : '-',
                            data ? data.pulang.pukul : '-',
                            data ? data.pulang.kurang : '-',

                            data ? (data.TL1 ? data.TL1 : '-') : '-',
                            data ? (data.TL2 ? data.TL2 : '-') : '-',
                            data ? (data.TL3 ? data.TL3 : '-') : '-',
                            data ? (data.TL4 ? data.TL4 : '-') : '-',

                            data ? (data.PSW1 ? data.PSW1 : '-') : '-',
                            data ? (data.PSW2 ? data.PSW2 : '-') : '-',
                            data ? (data.PSW3 ? data.PSW3 : '-') : '-',
                            data ? (data.PSW4 ? data.PSW4 : '-') : '-',
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

const getPresensi = (presensiArray, current_day) => {
    const isBefore1130 = presensiArray[0] ? moment(presensiArray[0], 'YYYY/MM/DD HH:mm:ss').isBefore(moment(current_day).hour(11).minute(29).second(59)) : undefined;
    if (presensiArray.length > 1) {
        const isAfter1330 = moment(presensiArray[presensiArray.length - 1], 'YYYY/MM/DD HH:mm:ss').isAfter(moment(current_day).hour(13).minute(29).second(59));
        return {
            datang: isBefore1130 ? moment(presensiArray[0], 'YYYY/MM/DD HH:mm:ss') : undefined,
            mid: (() => {
                let _mid = undefined;
                presensiArray.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isAfter(moment(current_day).hour(11).minute(29).second(59)) && moment(t, 'YYYY/MM/DD HH:mm:ss').isBefore(moment(current_day).hour(13).minute(29).second(59))) {
                        _mid = t
                    }
                })
                return _mid ? moment(_mid, 'YYYY/MM/DD HH:mm:ss') : _mid
            })(),
            pulang: isAfter1330 ? moment(presensiArray[presensiArray.length - 1], 'YYYY/MM/DD HH:mm:ss') : undefined,
        }
    } else {
        if (presensiArray[0]) {
            const isAfter1330 = moment(presensiArray[0], 'YYYY/MM/DD HH:mm:ss').isAfter(moment(current_day).hour(13).minute(29).second(59));
            return {
                datang: isBefore1130 ? moment(presensiArray[0], 'YYYY/MM/DD HH:mm:ss') : undefined,
                mid: !isBefore1130 && !isAfter1330 ? moment(presensiArray[0], 'YYYY/MM/DD HH:mm:ss') : undefined,
                pulang: isAfter1330 ? moment(presensiArray[0], 'YYYY/MM/DD HH:mm:ss') : undefined
            }
        } else {
            return {
                datang: undefined,
                mid: undefined,
                pulang: undefined
            }
        }
    }
}

const getPresensiShift = (presensi, current_day, name) => {
    if (!this.isShiftMalam(name)) return this.getPresensi(this.getAllDayHandkey(presensi).today, current_day)
    const isUp1800 = current_day.isAfter(moment(current_day).hour(17).minute(59).second(59))
    if (isUp1800) {
        return {
            datang: (() => {
                let _datang = undefined;
                this.getAllDayHandkey(presensi).today.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isBetween(
                        moment(current_day).hour(17).minute(59).second(59),
                        moment(current_day).hour(23).minute(29).second(59)
                    )) {
                        if (!_datang) _datang = t
                    }
                })
                return _datang ? moment(_datang, 'YYYY/MM/DD HH:mm:ss') : _datang
            })(),
            mid: (() => {
                let _mid = undefined;
                this.getAllDayHandkey(presensi).today.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isBetween(
                        moment(current_day).hour(23).minute(29).second(59),
                        moment(current_day).hour(23).minute(59).second(59)
                    )) {
                        _mid = t
                    }
                })
                return _mid ? moment(_mid, 'YYYY/MM/DD HH:mm:ss') : _mid
            })(),
            pulang: undefined
        }
    } else {
        return {
            datang: (() => {
                let _datang = undefined;
                this.getAllDayHandkey(presensi).yest.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isBetween(
                        moment(current_day).subtract(1, 'day').hour(17).minute(59).second(59),
                        moment(current_day).subtract(1, 'day').hour(23).minute(29).second(59)
                    )) {
                        _datang = t
                    }
                })
                return _datang ? moment(_datang, 'YYYY/MM/DD HH:mm:ss') : _datang
            })(),
            mid: (() => {
                let _mid = undefined;
                this.getAllDayHandkey(presensi).yest.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isBetween(
                        moment(current_day).subtract(1, 'day').hour(23).minute(29).second(59),
                        moment(current_day).hour(1).minute(30).second(0)
                    )) {
                        _mid = t
                    }
                })
                this.getAllDayHandkey(presensi).today.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isBetween(
                        moment(current_day).subtract(1, 'day').hour(23).minute(29).second(59),
                        moment(current_day).hour(1).minute(30).second(0)
                    )) {
                        _mid = t
                    }
                })
                return _mid ? moment(_mid, 'YYYY/MM/DD HH:mm:ss') : _mid
            })(),
            pulang: (() => {
                let _pulang = undefined;
                this.getAllDayHandkey(presensi).today.forEach(t => {
                    if (moment(t, 'YYYY/MM/DD HH:mm:ss').isBetween(
                        moment(current_day).hour(7).minute(29).second(59),
                        moment(current_day).hour(11).minute(30).second(0)
                    )) {
                        _pulang = t
                    }
                })
                return _pulang ? moment(_pulang, 'YYYY/MM/DD HH:mm:ss') : _pulang
            })()
        }
    }
}

const getAllDayHandkey = (presensi) => {
    return {
        today: presensi[0]._id === today_id ? presensi[0].handkey_time : presensi[1].handkey_time, // 'YYYY/MM/DD HH:mm:ss'
        yest: presensi[0]._id === yest_id ? presensi[0].handkey_time : presensi[1].handkey_time // 'YYYY/MM/DD HH:mm:ss'
    }
}