import React, { Component, Fragment } from 'react';
import ReactExport from 'react-data-export';
import { connect } from "react-redux";
import PropTypes from "prop-types";
import moment from "moment";

const ExcelFile = ReactExport.ExcelFile;
const ExcelSheet = ReactExport.ExcelFile.ExcelSheet;
const ExcelColumn = ReactExport.ExcelFile.ExcelColumn;
const ExcelCell = ReactExport.ExcelFile.ExcelCell;

const ExportExcel = ({
    shiftRegisters,
    monday,
    tuesday,
    wednesday,
    thursday,
    friday,
    saturday,
    sunday,
    users,
    shifts,
    jobs,
    branchName,
}) => {

    const multiDataSet = [
        {
            // Line Header
            columns: [
                { title: `${branchName}`, width: { wpx: 200 }, style: { font: { sz: "24", bold: true } } },
            ],
            data: [
                [
                    { value: "" },
                ]
            ]
        }
    ]

    let eleCols = [];
    let eleData = [];
    let tempArray = [];
    let shiftsLength = 0;
    shiftsLength = shifts.length;
    if (shiftsLength > 0) {
        // Line 2, 3
        eleCols.push(
            {
                title: "Họ Và Tên",
                width: { wch: 40 },
                style: {
                    font: { sz: "14", bold: false },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            }
        );
        eleData.push(
            {
                value: "",
                style: {
                    font: { sz: "14", bold: false },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            }
        );
        // Đếm theo số lượng ca hiện tại để tạo số ô đúng theo số ca
        let countShiftsLength = 0;
        // Lấy giá trị điểm ở giữa của ca. Ví dụ 3 ca => 3 ô => ô ở giữa sẽ điền thứ ngày vào đó
        let middleShiftsLength = 0;
        middleShiftsLength = shiftsLength / 2;
        let thu = 1;
        let setThu = "";
        let ngay = null;
        for (let i = 0; i < shiftsLength * 7; i++) {
            countShiftsLength = countShiftsLength + 1;
            if (countShiftsLength === 1) {
                thu = thu + 1;
                eleCols.push(
                    {
                        title: "",
                        width: { wpx: 50 },
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            },
                            fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                        }
                    }
                );
                eleData.push(
                    {
                        value: "",
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            },
                            fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                        }
                    }
                );
            }
            if (countShiftsLength !== 1 && countShiftsLength !== shiftsLength) {
                if (countShiftsLength === middleShiftsLength) {
                    if (thu === 2) {
                        setThu = "Thứ 2";
                        ngay = monday;
                    }
                    if (thu === 3) {
                        setThu = "Thứ 3";
                        ngay = tuesday;
                    }
                    if (thu === 4) {
                        setThu = "Thứ 4";
                        ngay = wednesday;
                    }
                    if (thu === 5) {
                        setThu = "Thứ 5";
                        ngay = thursday;
                    }
                    if (thu === 6) {
                        setThu = "Thứ 6";
                        ngay = friday;
                    }
                    if (thu === 7) {
                        setThu = "Thứ 7";
                        ngay = saturday;
                    }
                    if (thu === 8) {
                        setThu = "Chủ nhật";
                        ngay = sunday;
                    }
                    eleCols.push(
                        {
                            title: setThu,
                            width: { wpx: 50 },
                            style: {
                                font: { sz: "14", bold: false },
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                            }
                        }
                    );
                    eleData.push(
                        {
                            value: `${moment(ngay).format('MM/DD')}`,
                            style: {
                                font: { sz: "14", bold: false },
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                            }
                        }
                    );
                } else {
                    eleCols.push(
                        {
                            title: "",
                            width: { wpx: 50 },
                            style: {
                                font: { sz: "14", bold: false },
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                            }
                        }
                    );
                    eleData.push(
                        {
                            value: "",
                            style: {
                                font: { sz: "14", bold: false },
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                            }
                        }
                    );
                }

            }
            if (countShiftsLength === shiftsLength) {
                eleCols.push(
                    {
                        title: "",
                        width: { wpx: 50 },
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            },
                            fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                        }
                    }
                );
                eleData.push(
                    {
                        value: "",
                        style: {
                            font: { sz: "14", bold: false },
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            },
                            fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                        }
                    }
                );
                countShiftsLength = 0;
            }
        }
        // Add 3 columns lương
        eleCols.push(
            {
                title: "",
                width: { wpx: 50 },
                style: {
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            },
            {
                title: "Lương",
                width: { wpx: 50 },
                style: {
                    font: { sz: "14", bold: false },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            },
            {
                title: "",
                width: { wpx: 50 },
                style: {
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            }
        );
        eleData.push(
            {
                value: "",
                style: {
                    font: { sz: "14", bold: false },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            },
            {
                value: "",
                style: {
                    font: { sz: "14", bold: false },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            },
            {
                value: "",
                style: {
                    font: { sz: "14", bold: false },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } }
                    },
                    fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
                }
            }
        );
        tempArray.push(eleData);
        if (tempArray.length > 0) {
            multiDataSet.push({ columns: eleCols, data: tempArray });
        }
    }

    // const multiDataSet = [
    //     {
    //         // Line Header
    //         columns: [
    //             { title: `${branchName}`, width: { wpx: 200 }, style: { font: { sz: "24", bold: true } } },
    //         ],
    //         data: [
    //             [
    //                 { value: "" },
    //             ]
    //         ]
    //     },
    //     {
    //         // Line 2, 3
    //         columns: [
    //             {
    //                 title: "Họ Và Tên",
    //                 width: { wch: 40 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Thứ 2",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Thứ 3",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Thứ 4",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Thứ 5",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Thứ 6",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Thứ 7",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Chủ nhật",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         left: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "Lương",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     font: { sz: "14", bold: false },
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //             {
    //                 title: "",
    //                 width: { wpx: 50 },
    //                 style: {
    //                     border: {
    //                         top: { style: "thin", color: { rgb: "000000" } },
    //                         right: { style: "thin", color: { rgb: "000000" } },
    //                         bottom: { style: "thin", color: { rgb: "000000" } }
    //                     },
    //                     fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                 }
    //             },
    //         ],
    //         data: [
    //             [
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(monday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(tuesday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(wednesday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(thursday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(friday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(saturday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: `${moment(sunday).format('MM/DD')}`,
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             left: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //                 {
    //                     value: "",
    //                     style: {
    //                         font: { sz: "14", bold: false },
    //                         border: {
    //                             top: { style: "thin", color: { rgb: "000000" } },
    //                             right: { style: "thin", color: { rgb: "000000" } },
    //                             bottom: { style: "thin", color: { rgb: "000000" } }
    //                         },
    //                         fill: { patternType: "solid", fgColor: { rgb: "c7d6a1" } }
    //                     }
    //                 },
    //             ]
    //         ]
    //     }
    // ];

    eleCols = [];
    eleData = [];
    let getTotalColumns = shiftsLength * 7;
    let getIndexPosition = [];
    let getPosition = 0;
    const resetValue = (idx) => {
        for (let j = 0; j < shiftsLength; j++) {
            getIndexPosition.push(shifts.findIndex(x => x.position === j));
        }
        eleData = [];
        // eleCols = [];
        // for (let i = 0; i <= 24; i++) {
        for (let i = 0; i <= getTotalColumns + 3; i++) {
            // Set value default for Data
            // if (i < 22) {
            if (i < getTotalColumns + 1) {
                eleData.push({
                    value: "",
                    style: {
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        }
                    }
                });
            }
            // if (i === 22) {
            if (i === getTotalColumns + 1) {
                eleData.push({
                    value: "",
                    style: {
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        }
                    }
                });
            }
            if (i === getTotalColumns + 2) {
                eleData.push({
                    value: "",
                    style: {
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        }
                    }
                });
            }
            if (i === getTotalColumns + 3) {
                eleData.push({
                    value: "",
                    style: {
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        }
                    }
                });
            }
            // Set value default for Columns
            if (idx === 0) {
                if (i === 0) {
                    eleCols.push({
                        title: "",
                        width: { wch: 40 },
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            }
                        }
                    })
                }
                if (i < getTotalColumns + 1) {
                    // Ca 1
                    // if (i === 1 || i === 4 || i === 7 || i === 10 || i === 13 || i === 16 || i === 19) {
                    if (i > 0 && (i - 1) % shiftsLength === 0) {
                        getPosition = getIndexPosition[0];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "71f74a" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 2
                    // if (i === 2 || i === 5 || i === 8 || i === 11 || i === 14 || i === 17 || i === 20) {
                    if (i > 1 && (i - 2) % shiftsLength === 0) {
                        getPosition = getIndexPosition[1];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "4ba6ed" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 3
                    // if (i === 3 || i === 6 || i === 9 || i === 12 || i === 15 || i === 18 || i === 21) {
                    if (i > 2 && (i - 3) % shiftsLength === 0 && shiftsLength >= 3) {
                        getPosition = getIndexPosition[2];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "f5c344" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 4
                    if (i > 3 && (i - 4) % shiftsLength === 0 && shiftsLength >= 4) {
                        getPosition = getIndexPosition[3];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "7B68EE" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 5
                    if (i > 4 && (i - 5) % shiftsLength === 0 && shiftsLength >= 5) {
                        getPosition = getIndexPosition[4];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "8A2BE2" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 6
                    if (i > 5 && (i - 6) % shiftsLength === 0 && shiftsLength >= 6) {
                        getPosition = getIndexPosition[5];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "DDA0DD" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 7
                    if (i > 6 && (i - 7) % shiftsLength === 0 && shiftsLength >= 7) {
                        getPosition = getIndexPosition[6];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "D2B48C" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 8
                    if (i > 7 && (i - 8) % shiftsLength === 0 && shiftsLength >= 8) {
                        getPosition = getIndexPosition[7];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "3CB371" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 9
                    if (i > 8 && (i - 9) % shiftsLength === 0 && shiftsLength >= 9) {
                        getPosition = getIndexPosition[8];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "BDB76B" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                    // Ca 10
                    if (i > 9 && (i - 10) % shiftsLength === 0 && shiftsLength >= 10) {
                        getPosition = getIndexPosition[9];
                        eleCols.push({
                            title: `${shifts[getPosition].shiftName}`,
                            style: {
                                border: {
                                    top: { style: "thin", color: { rgb: "000000" } },
                                    left: { style: "thin", color: { rgb: "000000" } },
                                    right: { style: "thin", color: { rgb: "000000" } },
                                    bottom: { style: "thin", color: { rgb: "000000" } }
                                },
                                fill: { patternType: "solid", fgColor: { rgb: "7CFC00" } },
                                alignment: {
                                    horizontal: "center"
                                }
                            }
                        })
                    }
                }

                // if (i === 22) {
                if (i === getTotalColumns + 1) {
                    eleCols.push({
                        title: "",
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            }
                        }
                    })
                }
                // if (i === 23) {
                if (i === getTotalColumns + 2) {
                    eleCols.push({
                        title: "",
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            }
                        }
                    })
                }
                // if (i === 24) {
                if (i === getTotalColumns + 3) {
                    eleCols.push({
                        title: "",
                        style: {
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } }
                            }
                        }
                    })
                }
            }
        }
    }

    let getUsers = null;
    let shiftIndex = null;
    let jobIndex = null;
    let checkDayFlag = null;
    let costAmount = 0;
    tempArray = [];
    shiftRegisters.map((ele, idx) => {
        resetValue(idx);
        costAmount = 0;
        getUsers = users.find(({ _id }) => _id === ele.userId);
        // Add name
        eleData[0] =
        {
            value: `${getUsers.name}`,
            style: {
                font: { sz: "14", bold: false },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } }
                }
            }

        };
        ele.register.map((reg) => {
            shiftIndex = shifts.findIndex(x => x._id === reg.shiftId);
            jobIndex = jobs.findIndex(x => x._id === reg.jobId);
            costAmount = costAmount + reg.cost;
            // Add monday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(monday).format('MM-DD-YYYY')) {
                checkDayFlag = 1;
            }
            // Add tuesday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(tuesday).format('MM-DD-YYYY')) {
                checkDayFlag = shiftsLength + 1;
            }
            // Add wednesday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(wednesday).format('MM-DD-YYYY')) {
                checkDayFlag = shiftsLength * 2 + 1;
            }
            // Add thursday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(thursday).format('MM-DD-YYYY')) {
                checkDayFlag = shiftsLength * 3 + 1;
            }
            // Add friday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(friday).format('MM-DD-YYYY')) {
                checkDayFlag = shiftsLength * 4 + 1;
            }
            // Add saturday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(saturday).format('MM-DD-YYYY')) {
                checkDayFlag = shiftsLength * 5 + 1;
            }
            // Add sunday
            if (moment(reg.date).format('MM-DD-YYYY') === moment(sunday).format('MM-DD-YYYY')) {
                checkDayFlag = shiftsLength * 6 + 1;
            }

            if (shiftIndex === 0) {
                // Ca 1
                eleData[checkDayFlag] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "71f74a" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 1) {
                // Ca 2
                eleData[checkDayFlag + 1] =
                {
                    value: `${jobs.length >= 0 ? jobs[jobIndex].jobName : ""}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "4ba6ed" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 2) {
                // Ca 3
                eleData[checkDayFlag + 2] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "f5c344" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 3) {
                // Ca 4
                eleData[checkDayFlag + 3] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "7B68EE" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 4) {
                // Ca 5
                eleData[checkDayFlag + 4] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "8A2BE2" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 5) {
                // Ca 6
                eleData[checkDayFlag + 5] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "DDA0DD" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 6) {
                // Ca 7
                eleData[checkDayFlag + 6] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "D2B48C" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 7) {
                // Ca 8
                eleData[checkDayFlag + 7] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "3CB371" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 8) {
                // Ca 9
                eleData[checkDayFlag + 8] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "BDB76B" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
            if (shiftIndex === 9) {
                // Ca 10
                eleData[checkDayFlag + 9] =
                {
                    value: `${jobs[jobIndex].jobName}`,
                    style: {
                        font: { sz: "14", bold: false },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        },
                        fill: { patternType: "solid", fgColor: { rgb: "7CFC00" } },
                        alignment: {
                            horizontal: "center"
                        }
                    }

                };
            }
        })
        // Add salary of personal
        // eleData[24] =
        eleData[getTotalColumns + 3] = 
        {
            value: `${costAmount}`,
            style: {
                font: { sz: "15", bold: true },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } }
                },
                numFmt: "#,###",
                alignment: {
                    horizontal: "right"
                }
            }

        };
        tempArray.push(eleData);
    })
    if (tempArray.length > 0) {
        multiDataSet.push({ columns: eleCols, data: tempArray });
    }

    // multiDataSet.push({ columns: eleCols, data: [eleData] });
    // console.log("output " + JSON.stringify(multiDataSet));
    // console.log("output "+JSON.stringify(eleCols));
    // console.log("output "+JSON.stringify(eleData));
    const sheetName1 = moment(monday).format('MM-DD');
    const sheetName2 = moment(sunday).format('MM-DD');

    return (
        <Fragment>
            <div>
                <ExcelFile filename="Dang ki ca" element={<button type="button" class="btn btn-warning"
                    style={{ marginLeft: "10px" }}><i class="ti-import"></i>{"  "}Xuất Excel</button>}>
                    <ExcelSheet dataSet={multiDataSet} name={`${sheetName1} ~ ${sheetName2}`} />
                </ExcelFile>
            </div>
        </Fragment>
    );
};

ExportExcel.propTypes = {
    shiftRegisters: PropTypes.object.isRequired,
    monday: PropTypes.object.isRequired,
    tuesday: PropTypes.object.isRequired,
    wednesday: PropTypes.object.isRequired,
    thursday: PropTypes.object.isRequired,
    friday: PropTypes.object.isRequired,
    saturday: PropTypes.object.isRequired,
    sunday: PropTypes.object.isRequired,
    users: PropTypes.object.isRequired,
    shifts: PropTypes.object.isRequired,
    jobs: PropTypes.object.isRequired,
    branchName: PropTypes.object.isRequired,
};

export default connect(null, {})(
    ExportExcel
);