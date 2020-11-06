window.onload = () => {
    console.log('load');
    const fileInput = document.getElementById('fileInput');
    fileInput.addEventListener('change', fileHandler, false);

    function scheduleProcess(data) {
        const rawInfo = preProcessData(data);
        console.log(rawInfo);
    }

    function preProcessData(data) {
        const teachersInfo = parseTeacherInfo(data);
        const classInfo = parseClassInfo(data);
        const examInfo = parseExamInfo(data, classInfo);
        return {
            teachersInfo,
            examInfo
        };

        /**
         * parse 教師資訊
         * @param {*} data 
         */
        function parseTeacherInfo(data) {
            /*
             * schema
             * '教師姓名': {
             *      '姓名': '王力宏',
             *      '教學科目': '國文',
             *      '監考時數': 3,
             *      '進修時段': {
             *          '二': [1, 3],
             *          '四': [4, 5]
             *      }
             * }
             */
            const teachersInfo = {};
            if (data[0].name === '教師名單') {
                for (const key in data[0].rows) {
                    const row = data[0].rows[key].cells;
                    if (key === '0') {
                        // header
                        continue;
                    }
                    // schema: 0:姓名 1:教學科目 2:監考時數 3:進修時段
                    teachersInfo[row[0].text] = {
                        '姓名': row[0].text.trim(), //require
                        '教學科目': row[1] ? row[1].text.trim() : '',
                        '監考時數': +row[2].text, //require
                        '進修時段': row[3] ? parseStudyTime(row[3].text) : {}
                    }
                }
            }
            return teachersInfo;
            /**
             * parse進修時段
             * @param {*} rawTimeString 
             */
            function parseStudyTime(rawTimeString) {
                if (!rawTimeString) return {};
                // 三, 一(1), 二(2,4), 三(3-5)
                const times = {};
                rawTimeString.split(/(?<!()),(?!\))/i).forEach((dayString) => {
                    const day = dayString[0];
                    const timeString = dayString.substring(dayString.indexOf('(') + 1, dayString.indexOf(')'));
                    times[day] = [];
                    // 三 -> 全天
                    if (!timeString) {
                        for (i = 0; i < 5; i++) {
                            times[day].push(i + 1);
                        }
                        return true;
                    }
                    // 一(1)
                    if (timeString.length === 1) {
                        times[day].push(+timeString);
                        return true;
                    }
                    // 二(2,4)
                    if (timeString.includes(',')) {
                        timeString.split(',').forEach((time) => {
                            times[day].push(+time);
                        });
                        return true;
                    }
                    // 三(3-5)
                    if (timeString.includes('-')) {
                        const begin = +timeString.split('-')[0];
                        const end = +timeString.split('-')[1];
                        for (i = begin; i <= end; i++) {
                            times[day].push(i);
                        }
                        return true;
                    }
                    throw `timeString format error: ${timeString}`;
                });
                return times;
            }
        }
        /**
         * parse 考程資訊
         * @param {*} data 
         */
        function parseExamInfo(data, classInfo) {
            const examInfo = {};
            if (data[2].name === '考程') {
                for (const key in data[2].rows) {
                    const cells = data[2].rows[key].cells;
                    if (key === '0') {
                        // header
                        continue;
                    }
                    preprocessExamInfo(examInfo, cells);
                }
            }
            addClassInfoToExamInfo(examInfo, classInfo);
            return examInfo;

            function preprocessExamInfo(examInfo, cells) {
                const _day = cells[0].text;
                const _time = cells[1].text;
                const _grade = cells[2].text;
                const _group = cells[3].text;
                const _subject = cells[4].text;
                const day = examInfo[_day] || {};
                const time = day[_time] || {};
                const grade = time[_grade] || [];
                grade.push({
                    'subject': _subject,
                    'group': _group.split(','),
                    'classes': []
                });
                time[_grade] = grade;
                day[_time] = time;
                examInfo[_day] = day;
            }

            function addClassInfoToExamInfo(examInfo, classInfo) {
                for (const dayKey in examInfo) {
                    const day = examInfo[dayKey];
                    for (const timeKey in day) {
                        const time = day[timeKey];
                        for (const gradeKey in time) {
                            const grade = time[gradeKey];
                            grade.forEach(exam => {
                                exam.classes = findMatchClasses(gradeKey, exam.group);
                            });
                        }
                    }
                }

                function findMatchClasses(grade, groups) {
                    const matchClasses = [];
                    const gradeClasses = classInfo[grade];
                    for (const className in gradeClasses) {
                        const _class = gradeClasses[className];
                        if (groups.includes(_class['類組'])) {
                            matchClasses.push(_class);
                        }
                    }
                    return matchClasses;
                }
            }
        }
        /**
         * parse 班級資訊
         * @param {*} data 
         */
        function parseClassInfo(data) {
            const classInfo = {};
            if (data[1].name === '班級') {
                for (const key in data[1].rows) {
                    const row = data[1].rows[key].cells;
                    if (key === '0') {
                        // header
                        continue;
                    }
                    const grade = classInfo[row[0].text] ? classInfo[row[0].text] : {};
                    grade[row[1].text.trim()] = {
                        '樓層': +row[2].text,
                        '類組': row[3].text.trim()
                    };
                    classInfo[row[0].text] = grade;
                }
            }
            return classInfo;
        }
    }

    function workbookHandler(workbook) {
        // console.log('handle workbook');
        const jsonObject = [];
        workbook.SheetNames.forEach(sheetName => {
            const parsedSheet = {
                name: sheetName,
                header: {},
                rows: {}
            };
            const worksheet = workbook.Sheets[sheetName];
            const aoa = XLSX.utils.sheet_to_json(worksheet, {
                raw: false,
                header: 1
            });
            aoa.forEach((row, rowIndex) => {
                const cells = {};
                row.forEach((cell, cellIndex) => {
                    cells[cellIndex] = ({
                        text: cell
                    });
                });
                parsedSheet.rows[rowIndex] = {
                    cells: cells
                };
                // header
                if (rowIndex === 0) {
                    parsedSheet.header = cells;
                }
            });
            jsonObject.push(parsedSheet);
        });
        return jsonObject;
    }

    function fileReader(event) {
        const rawData = new Uint8Array(event.target.result);
        var workbook = XLSX.read(rawData, {
            type: 'array'
        });
        const parsedJson = workbookHandler(workbook);
        scheduleProcess(parsedJson);
        // reset file input
        fileInput.value = '';
    }

    function fileHandler(event) {
        const rawfile = event.target.files[0];
        if (rawfile) {
            // console.log('dealing with file');
            const reader = new FileReader();
            reader.onload = fileReader;
            reader.readAsArrayBuffer(rawfile);
        }
    }
}