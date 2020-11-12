window.onload = () => {
    console.log('load');
    const fileInput = document.getElementById('fileInput');
    fileInput.addEventListener('change', fileHandler, false);

    const setTools = {
        union: (set1, set2) => {
            const _union = new Set(set1);
            for (const element of set2) {
                _union.add(element);
            }
            return _union;
        }
    }
    const timeTools = {
        parseTime: (rawTimeString) => {
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
        },
        mergeTime: (time1, time2) => {
            const time = {};
            const intersection = [];
            Object.assign(time, time1);
            for (const dayInTime2 in time2) {
                if (dayInTime2 in time) {
                    intersection.push(dayInTime2);
                    const set1 = new Set(time[dayInTime2]);
                    const set2 = time2[dayInTime2];
                    const _union = setTools.union(set1, set2);
                    time[dayInTime2] = [..._union];
                    continue;
                }
                time[dayInTime2] = time2[dayInTime2];
            }
            // for (const dayInTim1)
            return time;
        }
    }

    function getRandomInt(max) {
        return Math.floor(Math.random() * Math.floor(max));
    }

    function getRandomNumbers(ammount, min, max) {
        const range = max - min + 1;
        if (range < ammount) return [];
        const arr = [];
        while (arr.length < ammount) {
            const tmpNumber = getRandomInt(range);
            if (!arr.includes(tmpNumber)) {
                arr.push(tmpNumber);
            }
        }
        return arr.map(element => element + min);
    }

    function scheduleProcess(data) {
        function constainsDayTime(unavailableTime, day, time) {
            const times = {
                '第一節': 1,
                '第二節': 2,
                '第三節': 3,
                '第四節': 4,
                '第五節': 5,
            };
            const unavailableDays = Object.keys(unavailableTime);
            // 這天沒不行
            if (!unavailableDays.includes(day)) return false;
            return unavailableTime[day].includes(times[time]);
        }

        function getAvailableTeachers(teachersInfo, day, time) {
            const names = [];
            for (name in teachersInfo) {
                if (constainsDayTime(teachersInfo[name]['排除時段'], day, time)) {
                    continue;
                }
                if (!teachersInfo[name]['監考時數']) {
                    continue;
                }
                names.push(name);
            }
            return names;
        }

        function updateTeacherInfo(teachersInfo, name, day, time, _class) {
            teachersInfo[name]['監考時數'] -= 1;
            const teacherExam = teachersInfo[name]['監考'] || [];
            teacherExam.push({
                day,
                time,
                _class
            });
            teachersInfo[name]['監考'] = teacherExam;
        }
        let rawInfo = preProcessData(data);
        // const backup = Object.assign({}, rawInfo);
        const backup = JSON.stringify(rawInfo);
        const maxTry = 99;
        let passFlag = true;
        let tryCount = 0;
        console.log(rawInfo);
        // const test1 = getAvailableTeachers(rawInfo.teachersInfo, '三', '第五節');
        // const test2 = getAvailableTeachers(rawInfo.teachersInfo, '五', '第三節');
        // console.log({
        //     test1,
        //     test2
        // });
        do {
            passFlag = true;
            rawInfo = JSON.parse(backup);
            console.log(`%c 第${tryCount + 1}次嘗試`, 'color: #e4cdb0');
            for (const day in rawInfo.examInfo) {
                for (const time in rawInfo.examInfo[day]) {
                    let teacherNeededAmount = 0;
                    for (const grade in rawInfo.examInfo[day][time]) {
                        teacherNeededAmount += rawInfo.examInfo[day][time][grade].length
                    }
                    const availableTeacherList = getAvailableTeachers(rawInfo.teachersInfo, day, time);
                    if (teacherNeededAmount > availableTeacherList.length) {
                        console.log(`%c 星期${day}, ${time} 監考老師不足,需要${teacherNeededAmount},僅有${availableTeacherList.length}人 `, 'color: #629dd0');
                        passFlag = false;
                        break;
                    }
                    const teacherIndex = getRandomNumbers(teacherNeededAmount, 0, availableTeacherList.length - 1);
                    const teachers = teacherIndex.map(i => availableTeacherList[i]);
                    // console.log(` day: ${day}, time: ${time} `);
                    // console.log({
                    //     teachers
                    // });
                    let i = 0;
                    for (const grade in rawInfo.examInfo[day][time]) {
                        rawInfo.examInfo[day][time][grade] = rawInfo.examInfo[day][time][grade].map(_class => {
                            const teacher = teachers[i++]; //  teachers.pop();
                            updateTeacherInfo(rawInfo.teachersInfo, teacher, day, time, _class);
                            return {
                                _class: _class,
                                teacher: teacher
                            };
                        });
                    }
                }
            }
            tryCount += 1;
        } while (!passFlag && tryCount < maxTry)
        if (!passFlag) {
            console.log(`%c 失敗了`, 'color: #e4cdb0');
        } else {
            console.log(`%c 成功了!!!!`, 'color: #e4cdb0');
        }
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
            function parseUnavailableTime(time1, time2) {
                const time1Text = time1 ? time1.text : '';
                const time2Text = time2 ? time2.text : '';
                if (time1Text && time2Text) {
                    return timeTools.mergeTime(timeTools.parseTime(time1Text), timeTools.parseTime(time2Text));
                }
                if (time1Text) {
                    return timeTools.parseTime(time1Text);
                }
                return timeTools.parseTime(time2Text);
            }
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
                    const name = row[0].text.trim();
                    teachersInfo[name] = {
                        '姓名': name, //require
                        '教學科目': row[1] ? row[1].text.trim() : '',
                        '監考時數': +row[2].text, //require
                        '排除時段': parseUnavailableTime(row[3], row[4])
                    }
                }
            }
            return teachersInfo;
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
            // addClassInfoToExamInfo(examInfo, classInfo);
            return examInfo;

            function parseClass(_class) {
                const c = [];
                _class.split(',').forEach(range => {
                    const splitResult = range.split('-');
                    if (splitResult.length === 2) {
                        for (let i = +splitResult[0]; i <= +splitResult[1]; i++) {
                            c.push(i);
                        }
                        return true;
                    }
                    c.push(+splitResult[0]);
                });
                return c;
            }

            function preprocessExamInfo(examInfo, cells) {
                const _day = cells[0].text;
                const _time = cells[1].text;
                const _grade = cells[2].text;
                const _class = parseClass(cells[3].text);
                const day = examInfo[_day] || {};
                const time = day[_time] || {};
                const grade = time[_grade] || [];
                time[_grade] = _class;
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
                            matchClasses.push({
                                'className': className,
                                '類組': _class['類組'],
                                '樓層': _class['樓層'],
                            });
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
                header: 1,
                blankrows: false
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