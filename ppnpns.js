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
let targetDay = moment();
const today_id = moment().format('YYYY_MM_DD')
const yest_id = moment().subtract(1, 'day').format('YYYY_MM_DD')
const formatTanggalWithHour = 'HH:mm:ss M/D/YYYY'
const formatHour = 'HH:mm:ss'
const formatTglRawData = 'M/D/YYYY';

const terlambat_menitF = (absen_time, jam_batas) => (absen_time.diff(jam_batas, 'minutes') > 0 ? absen_time.diff(jam_batas, 'minutes') : '-');
const psw_menitF = (absen_time, jam_batas) => (jam_batas.diff(absen_time, 'minutes') > 0 ? jam_batas.diff(absen_time, 'minutes') : '-');

let ppnpns = {}
data[0].data.forEach((row, i, arr) => {
    if (i > 0) {
        if (!groups[row[2]]) {
            groups[row[2]] = {
                absen: {}
            };
        }

        if (!ppnpns[row[2]]) ppnpns[row[2]] = {} //nama
        if (!ppnpns[row[2]]['absen']) ppnpns[row[2]]['absen'] = {} //absen obj
        if (!ppnpns[row[2]]['absen'][row[4]]) ppnpns[row[2]]['absen'][row[4]] = {} //tanggal
        if (!ppnpns[row[2]]['absen'][row[4]].datang) ppnpns[row[2]]['absen'][row[4]].datang = {
            pukul: undefined,
            telat: 999
        }
        if (!ppnpns[row[2]]['absen'][row[4]].mid) ppnpns[row[2]]['absen'][row[4]].mid = {
            pukul: undefined
        }
        if (!ppnpns[row[2]]['absen'][row[4]].pulang) ppnpns[row[2]]['absen'][row[4]].pulang = {
            pukul: undefined,
            kurang: 999
        }
        if (!ppnpns[row[2]]['absen'][row[4]].all_absen) ppnpns[row[2]]['absen'][row[4]].all_absen = []
        const active_ppnpns_absen = ppnpns[row[2]]['absen']
        const active_absen_today = ppnpns[row[2]]['absen'][row[4]]
        current_day = moment(row[4], formatTglRawData);
        if( i===1 ) targetDay = moment(row[4], formatTglRawData);
        yesterday = moment(row[4], formatTglRawData).subtract(1, 'day');
        besok = moment(row[4], formatTglRawData).add(1, 'day');
        const all_absen_yest = ppnpns[row[2]]['absen'][yesterday.format(formatTglRawData)] ? ppnpns[row[2]]['absen'][yesterday.format(formatTglRawData)].all_absen : []
        for (let index = 5; index <= 8; index++) {
            if (row[index]) {
                const absen_time = moment(`${row[index]} ${row[4]}`, formatTanggalWithHour)
                active_absen_today.all_absen.push(absen_time)
                const pukul0000 = moment(current_day).hour(0).minute(0).second(0)
                const pukul0130 = moment(current_day).hour(1).minute(29).second(59)
                const pukul0600 = moment(current_day).hour(5).minute(59).second(59)
                const pukul0730 = moment(current_day).hour(7).minute(29).second(59)
                const pukul1130 = moment(current_day).hour(11).minute(29).second(59)
                const pukul1330 = moment(current_day).hour(13).minute(29).second(59)
                const pukul1600 = moment(current_day).hour(15).minute(59).second(59)
                const pukul1630 = moment(current_day).hour(16).minute(29).second(59)
                const pukul1930 = moment(current_day).hour(19).minute(29).second(59)
                const pukul2330 = moment(current_day).hour(23).minute(29).second(59)
                const pukul2359 = moment(current_day).hour(23).minute(59).second(59)
                const pukul2330kemarin = moment(current_day).subtract(1, 'day').hour(23).minute(29).second(59)
                const tipe_pnpns = shift_ppnpn[row[2]][current_day.day()][2] // tipe1 atau tipe2
                if (shift_ppnpn[row[2]][current_day.day()][0]) { //SHIFT SIANG
                    if (absen_time.isBetween(pukul0600, pukul1130) && !active_absen_today.datang.pukul) { //jika antara 06.00 - 11.30 dianggap absen datang
                        active_absen_today.datang = {
                            pukul: absen_time,
                            telat: terlambat_menitF(absen_time, pukul0730)
                        }
                    } else if (absen_time.isBetween(pukul1130, pukul1330)) { //jika antara 11.30 - 13.30 dianggap absen mid
                        active_absen_today.mid = {
                            pukul: absen_time
                        }
                    } else if (absen_time.isBetween(pukul1130, pukul2330)) { //jika antara 11.30 - 23.30 dianggap absen pulang
                        active_absen_today.pulang = {
                            pukul: absen_time,
                            kurang: psw_menitF(absen_time, tipe_pnpns === 'tipe2' ? (absen_time.day() === 5 ? pukul1630 : pukul1600) : pukul1930)
                        }
                    }
                }
                const tgl_kemarin = yesterday.format(formatTglRawData)
                if (shift_ppnpn[row[2]][yesterday.day()][1]) { //KEMARIN SHIFT MALAM
                    if (!active_ppnpns_absen[tgl_kemarin]) active_ppnpns_absen[tgl_kemarin] = {} //tanggal
                    if (!active_ppnpns_absen[tgl_kemarin].datang) active_ppnpns_absen[tgl_kemarin].datang = {
                        pukul: undefined,
                        telat: 999
                    }
                    if (!active_ppnpns_absen[tgl_kemarin].mid) active_ppnpns_absen[tgl_kemarin].mid = {
                        pukul: undefined
                    }
                    if (!active_ppnpns_absen[tgl_kemarin].pulang) active_ppnpns_absen[tgl_kemarin].pulang = {
                        pukul: undefined,
                        kurang: 999
                    }
                    if (absen_time.isBetween(pukul0000, pukul0130)) { //jika antara 00.00 - 01.30 dianggap absen pulang
                        active_ppnpns_absen[tgl_kemarin].mid = {
                            pukul: absen_time
                        }
                    }
                    else if (absen_time.isBetween(pukul0130, pukul1130)) { //jika antara 00.00 - 01.30 dianggap absen pulang
                        active_ppnpns_absen[tgl_kemarin].pulang = {
                            pukul: absen_time,
                            kurang: psw_menitF(absen_time, pukul0730)
                        }
                    }
                }
                if (shift_ppnpn[row[2]][current_day.day()][1]) { //SHIFT MALAM
                    if (!active_ppnpns_absen[tgl_kemarin]) active_ppnpns_absen[tgl_kemarin] = {} //tanggal
                    if (!active_ppnpns_absen[tgl_kemarin].datang) active_ppnpns_absen[tgl_kemarin].datang = {
                        pukul: undefined,
                        telat: 999
                    }
                    if (!active_ppnpns_absen[tgl_kemarin].mid) active_ppnpns_absen[tgl_kemarin].mid = {
                        pukul: undefined
                    }
                    if (!active_ppnpns_absen[tgl_kemarin].pulang) active_ppnpns_absen[tgl_kemarin].pulang = {
                        pukul: undefined,
                        kurang: 999
                    }
                    if (absen_time.isBetween(pukul1600, pukul2330)) { //jika antara 16.00 - 23.30 dianggap absen datang
                        if (!active_absen_today.datang.pukul) {
                            active_absen_today.datang = {
                                pukul: absen_time,
                                telat: terlambat_menitF(absen_time, pukul1930)
                            }
                        }
                    }
                    else if (absen_time.isBetween(pukul2330, pukul2359)) { //jika antara 23.30 - 23.59 dianggap absen mid
                        active_absen_today.mid = {
                            pukul: absen_time
                        }
                    }
                    // else if (absen_time.isBetween(pukul0000, pukul0130)) { //jika antara 00.00 - 01.30 dianggap absen pulang
                    //     active_ppnpns_absen[tgl_kemarin].mid = {
                    //         pukul: absen_time
                    //     }
                    // }
                    // else if (absen_time.isBetween(pukul0130, pukul1130) && ) { //jika antara 00.00 - 01.30 dianggap absen pulang
                    //     active_ppnpns_absen[tgl_kemarin].pulang = {
                    //         pukul: absen_time,
                    //         kurang: psw_menitF(absen_time, pukul0730)
                    //     }
                    // }
                }

            }
        }
    }
})


XlsxPopulate.fromFileAsync(__dirname + "/rekap_ppnpns.xlsx")
    .then(workbook => {
        let index = 0;
        for (let nama in ppnpns) {
            if (ppnpns.hasOwnProperty(nama)) {
                if (true) {
                    let sheet = workbook.sheet(index);
                    sheet.name(nama)
                    workbook.sheet(index).cell("B1").value(targetDay.format('MMMM YYYY'));
                    workbook.sheet(index).cell("B2").value(nama);
                    let row = 7
                    for (let i = 1; i <= targetDay.endOf('month').date(); i++) {
                        let r = sheet.range('A' + row + ':G' + row);
                        (moment(targetDay).date(i).day() === 0 || moment(targetDay).date(i).day() === 6) && r.style("fill", {
                            type: "solid",
                            color: {
                                rgb: "8c8c8c"
                            }
                        })
                        let data = ppnpns[nama].absen[moment(targetDay).date(i).format(formatTglRawData)];
                        let arr = [
                            moment(targetDay).date(i).format('DD/MM/YYYY'),
                            moment(targetDay).date(i).format('dddd'),
                            data ? (data.datang.pukul ? data.datang.pukul.format(formatHour) : '-') : '-',
                            data ? (data.datang.telat ? data.datang.telat : '-') : '-',
                            data ? (data.mid.pukul ? data.mid.pukul.format(formatHour) : '-') : '-',
                            data ? (data.pulang.pukul ? data.pulang.pukul.format(formatHour) : '-') : '-',
                            data ? (data.pulang.kurang ? data.pulang.kurang : '-') : '-',

                            // data ? (data.TL1 ? data.TL1 : '-') : '-',
                            // data ? (data.TL2 ? data.TL2 : '-') : '-',
                            // data ? (data.TL3 ? data.TL3 : '-') : '-',
                            // data ? (data.TL4 ? data.TL4 : '-') : '-',

                            // data ? (data.PSW1 ? data.PSW1 : '-') : '-',
                            // data ? (data.PSW2 ? data.PSW2 : '-') : '-',
                            // data ? (data.PSW3 ? data.PSW3 : '-') : '-',
                            // data ? (data.PSW4 ? data.PSW4 : '-') : '-',
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

        if (fs.existsSync(__dirname + `/rekap_ok_PPNPNS.xlsx`)) {
            fs.unlinkSync(__dirname + `/rekap_ok_PPNPNS.xlsx`);
        }
        workbook.toFileAsync(__dirname + `/rekap_ok_PPNPNS.xlsx`);
    }).then(dataa => {
        console.log('Finished');
    })