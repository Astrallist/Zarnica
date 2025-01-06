document.getElementById('processButton').addEventListener('click', async () => {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert("Пожалуйста, выберите файл.");
        return;
    }

    const arrayBuffer = await file.arrayBuffer();

    // Извлечение текста из документа
    const { value } = await mammoth.extractRawText({ arrayBuffer });
    const oldName = file.name;
    const name = oldName.slice(0, -5);

    console.log("Извлеченный текст:", value);  // Для отладки

    // Обработка текста
    const tasks = parseTasks(value);
    
    console.log("Задания:", tasks);  // Для отладки

    // Проверка, содержит ли tasks данные
    if (tasks.length === 0) {
        alert("Не было найдено заданий для обработки.");
        return;
    }

    // Создание нового документа
    const doc = new docx.Document({
        sections: [{
            properties: {},
            children: tasks
        }]
    });

    // Генерация файла
    docx.Packer.toBlob(doc).then(blob => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `${name}_2.0.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    });

    alert("Получилось");
});

// Функция для разбора заданий
function parseTasks(text) {
    const tasks = [];
    const lines = text.split('\n');

    let taskContent = {
        title: "",
        team: "",
        squad: "",
        units: "",
        equipment: "",
        stitching: "",
        time: "",
        text: [],
        tasks: "",
        check: "",
        notes: ""
    };

    for (let line of lines) {
        line = line.trim();

        if (line.startsWith("Вводная:")) {
            if (taskContent.title) {
                const ctp = createTaskParagraph(taskContent);
                for (let i = 0; i<ctp.length;i++) {
                    tasks.push(ctp[i]);
                }
                taskContent = { title: "",
                    team: "",
                    squad: "",
                    units: "",
                    equipment: "",
                    stitching: "",
                    time: "",
                    text: [],
                    tasks: "",
                    check: "",
                    notes: "" };
            }
            taskContent.title = line;
        } else if (line.startsWith("Взвод:")) {
            taskContent.team = line;
        } else if (line.startsWith("Отделение:")) {
            taskContent.squad = line;
        } else if (line.startsWith("Силы:")) {
            taskContent.units = line;
        } else if (line.startsWith("Тех.средства:")) {
            taskContent.equipment = line;
        } else if (line.startsWith("Пришивка:")) {
            taskContent.stitching = line;
        } else if (line.startsWith("Время:")) {
            taskContent.time = line;
        } else if (line.startsWith("Текст:")) {
            const oldText = line;
            const text = oldText.slice(6);
            taskContent.text.push(text);
        } else if (line.startsWith("Задачи по пунктам посреднику:")) {
            taskContent.tasks = line;
        } else if (line.startsWith("Проверка и оценивание посреднику:")) {
            taskContent.check = line;
        } else if (line.startsWith("Примечания посреднику:")) {
            taskContent.notes = line;
        }
    }

    // Добавляем последнее задание если оно существует
    if (taskContent.title) {
        const ctp = createTaskParagraph(taskContent);
                for (let i = 0; i<ctp.length;i++) {
                    tasks.push(ctp[i]);
                };
    }

    console.log("Разобранные задания:", tasks); // Для отладки
    return tasks;
}

// Функция для создания абзацев
function createTaskParagraph(task) {
    console.log(task);

/*
        title: "",
        team: "",
        squad: "",
        units: "",
        equipment: "",
        stitching: "",
        time: "",
        text: "",
        tasks: "",
        check: "",
        notes: ""
        */

        let result = [
            new docx.Paragraph({
                //heading: docx.HeadingLevel.HEADING_2,
                alignment: docx.AlignmentType.CENTER,
                bold: true,
                children: [
                    new docx.TextRun({
                        text: task.title,
                        bold: true,
                    })
                ]
            }),
            new docx.Paragraph({
                text: task.team,
                alignment: docx.AlignmentType.LEFT,
            }),
            new docx.Paragraph({
                text: task.squad,
                alignment: docx.AlignmentType.LEFT,
            }),
            new docx.Paragraph({
                text: task.units,
                alignment: docx.AlignmentType.LEFT,
            }),
            new docx.Paragraph({
                text: task.equipment,
                alignment: docx.AlignmentType.LEFT,
            }),
            new docx.Paragraph({
                text: task.stitching,
                alignment: docx.AlignmentType.LEFT,
            }),
            new docx.Paragraph({
                text: task.time,
                alignment: docx.AlignmentType.LEFT,
            }),
        ];

        for (let i =0;i<task.text.length;i++) {
            result.push(        
                new docx.Paragraph({
                text: `     ${task.text[i]}`,
                alignment: docx.AlignmentType.LEFT,
                spacing: {before: 200},
                outlineLevel: 0,
            }),
        )}

        let result2 = result.concat([
            new docx.Paragraph("", { spacing: { before: 200 }}),
    
            new docx.Paragraph({
                alignment: docx.AlignmentType.justified,
                children: [
                    new docx.TextRun("Начальник штаба игры  "),
                    new docx.TextRun("                                                                                          ____________/Корнеева Е.И",),
                ],
            }),

            new docx.Paragraph("", {thematicBreak: true}),
            new docx.Paragraph("_________________________________________________________________________________________"),
            //ВТОРОЙ

            
            new docx.Paragraph({
                //heading: docx.HeadingLevel.HEADING_2,
                alignment: docx.AlignmentType.CENTER,
                children: [
                    new docx.TextRun({
                        text: task.title,
                        bold: true,
                        size: 20,
                    })
                ]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.LEFT,
                children: [
                    new docx.TextRun({
                        text: task.team,
                        size: 20,
                    })
                ]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.LEFT,
                children: [
                    new docx.TextRun({
                        text: task.squad,
                        size: 20,
                    })
                ]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.LEFT,
                children: [
                    new docx.TextRun({
                        text: task.units,
                        size: 20,
                    })
                ]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.LEFT,
                children: [
                    new docx.TextRun({
                        text: task.equipment,
                        size: 20,
                    })
                ]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.LEFT,
                children: [
                    new docx.TextRun({
                        text: task.stitching,
                        size: 20,
                    })
                ]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.LEFT,
                children: [
                    new docx.TextRun({
                        text: task.time,
                        size: 20,
                    })
                ]
            }),
            
            
            
            
            
            
            
            
            new docx.Paragraph("", { pageBreakAfter: true, pageBreak: true }), // Разрыв страницы
            new docx.PageBreak(),
        ])

        console.log(result2)

    return result2;
}
