// Require library
var excel = require('excel4node');
const random = require('random')

// Create a new instance of a Workbook class
var workbook = new excel.Workbook();

// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Sheet 1');

var cau_1 = ['Male', 'Female', 'Other'];
var cau_2 = ['18-20', '21-25', '25-30', 'Older than 30'];
var cau_3 = ['Totally agree', 'Agree', 'Disagree', 'Totally disagree'];
var cau_4 = ['Sure (please answer question 5)', 'Not sure (skip to question 6)'];
var cau_5 = [16, 17, 18, 19, 20, 21];
var cau_6 = ['Totally agree', 'Agree', 'Disagree', 'Totally disagree'];
var cau_7 = ['Totally aware', 'Partly aware', 'Not at all'];
var cau_8 = ['Poverty', 'Lack of awareness of child labour', 'Toxic parents', 'Loose laws and regulations', 'Disadvantaged children', 'Orphans', 'Child manipulators'];
var cau_9 = ['Children health', 'Shortage of accessing to education and basic development', 'Lack of chance for good job in the future', 'Physical abuse', 'Sexual abuse',
    'Emotional neglect (lonliness, hopelessness…)', 'Be exposed to harmful behaviors (smoking, drug use,…)'];
var cau_10 = ['Improving the legal system to protect child labors', 'Raise awareness of child labour', 'Provide immediate assistance to the victims',
    'Establish child abuse prevention organizations', 'Facilitating school fees for poor families', 'Provide employment opportunities for parents',
    'Increase scrutiny on suspected cases of child labour', 'Encourage people to report cases of child labour'];

// Create a reusable style
var style = workbook.createStyle({
    font: {
        size: 12
    }
});
var style_header = workbook.createStyle({
    font: {
        size: 13
    }
});

for (let i = 1; i <= 10; i++) {
    worksheet.cell(1, i).string('Câu ' + i).style(style_header);
}
for (let j = 2; j < 10003; j++) {
    for (let i = 1; i <= 10; i++) {
        let ran_index;
        let lst_ran_index = [];
        let string_random = '';
        switch (i) {
            case 1:
                ran_index = random.int((min = 0), (max = 1));
                string_random = cau_1[ran_index];
                break;
            case 2:
                ran_index = random.int((min = 0), (max = 3));
                string_random = cau_2[ran_index];
                break;
            case 3:
                ran_index = random.int((min = 0), (max = 3));
                string_random = cau_3[ran_index];
                break;
            case 4:
                ran_index = random.int((min = 0), (max = 1));
                string_random = cau_4[ran_index];
                break;
            case 5:
                ran_index = random.int((min = 0), (max = 5));
                break;
            case 6:
                ran_index = random.int((min = 0), (max = 3));
                string_random = cau_6[ran_index];
                break;
            case 7:
                ran_index = random.int((min = 0), (max = 2));
                string_random = cau_7[ran_index];
                break;
            case 8:
                ran_index = random.int((min = 2), (max = 4));
                for (let k = 0; k < ran_index; k++) {
                    let ran_item = random.int((min = 0), (max = 6));
                    var check = lst_ran_index.includes(ran_item);
                    while (check) {
                        ran_item = random.int((min = 0), (max = 6));
                        var check = lst_ran_index.includes(ran_item);
                    }
                    lst_ran_index.push(ran_item);
                    string_random += (string_random != '' ? (", " + cau_8[ran_index]) : cau_8[ran_index]);
                }
                break;
            case 9:
                ran_index = random.int((min = 2), (max = 4));
                for (let k = 0; k < ran_index; k++) {
                    let ran_item = random.int((min = 0), (max = 6));
                    var check = lst_ran_index.includes(ran_item);
                    while (check) {
                        ran_item = random.int((min = 0), (max = 6));
                        var check = lst_ran_index.includes(ran_item);
                    }
                    lst_ran_index.push(ran_item);
                    string_random += (string_random != '' ? (", " + cau_9[ran_index]) : cau_9[ran_index]);
                }
                break;
            case 10:
                ran_index = random.int((min = 2), (max = 4));
                for (let k = 0; k < ran_index; k++) {
                    let ran_item = random.int((min = 0), (max = 7));
                    var check = lst_ran_index.includes(ran_item);
                    while (check) {
                        ran_item = random.int((min = 0), (max = 7));
                        var check = lst_ran_index.includes(ran_item);
                    }
                    lst_ran_index.push(ran_item);
                    string_random += (string_random != '' ? (", " + cau_10[ran_index]) : cau_10[ran_index]);
                }
                break;
        }
        console.log("Câu " + i);
        console.log(string_random);
        if (i == 5)
            console.log(cau_5[ran_index]);
        if (i == 5)
            worksheet.cell(j, i).number(cau_5[ran_index]).style(style);
        else
            worksheet.cell(j, i).string(string_random).style(style);
    }
}

workbook.write('Excel.xlsx');