window.onload = () => {
    console.log('load');
    const fileInput = document.getElementById('fileInput');
    fileInput.addEventListener('change', fileHandler, false);

    function scheduleProcess(data) {
        const rawInfo = preProcessData(data);
        console.log(data);
    }

    function preProcessData(data) {
        const teachersInfo = parseTeacherInfo(data);
        const classInfo = parseClassInfo(data);
        const examInfo = parseExamInfo(data, classInfo);
        return {
            teachersInfo: teachersInfo
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
                Object.keys(data[0].rows).forEach((key) => {
                    const row = data[0].rows[key].cells;
                    if (key === '0') {
                        // header
                        // console.log(`header: {row}`);
                        return true;
                    }
                    // schema: 0:姓名 1:教學科目 2:監考時數 3:進修時段
                    teachersInfo[row[0].text] = {
                        '姓名': row[0].text.trim(), //require
                        '教學科目': row[1] ? row[1].text.trim() : '',
                        '監考時數': +row[2].text, //require
                        '進修時段': row[3] ? parseStudyTime(row[3].text) : {}
                    }
                });
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
                Object.keys(data[2].rows).forEach((key) => {
                    const row = data[2].rows[key].cells;
                    if (key === '0') {
                        // header
                        return true;
                    }
                    const day = examInfo[row[0].text] ? examInfo[row[0].text] : {};
                    const time = day[row[1].text] ? day[row[1].text] : {};
                    const grade = time[row[2].text] ? time[row[2].text] : {};
                    grade[]
                });
            }
            return examInfo;
        }
        /**
         * parse 班級資訊
         * @param {*} data 
         */
        function parseClassInfo(data) {
            const classInfo = {};
            if (data[1].name === '班級') {
                Object.keys(data[1].rows).forEach((key) => {
                    const row = data[1].rows[key].cells;
                    if (key === '0') {
                        // header
                        return true;
                    }
                    const grade = classInfo[row[0].text] ? classInfo[row[0].text] : {};
                    grade[row[1].text.trim()] = {
                        '樓層': +row[2].text,
                        '類組': row[3].text.trim()
                    };
                    classInfo[row[0].text] = grade;
                });
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