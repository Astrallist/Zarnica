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
            const oldCheck = line;
            const check = oldCheck.slice(33);
            taskContent.check = check;
        } else if (line.startsWith("Примечания посреднику:")) {
            const oldNotes = line;
            const notes = oldNotes.slice(22);
            taskContent.notes = notes;
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
                        size: 24,
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
                children: [
                    new docx.TextRun({
                        text: `     ${task.text[i]}`,
                        size: 24,
                    }),
                ],
                alignment: docx.AlignmentType.LEFT,
                spacing: {before: 200},
                outlineLevel: 0,
            }),
        )}

        let result2 = result.concat([
            new docx.Paragraph("", { spacing: { before: 200 }}),
    
            /*
            new docx.Table({
                borders: {
                    BorderOptions: {
                        top: {
                            size: 0,
                            color: "ffffff",
                        },
                        bottom: {
                            size: 0,
                            color: "ffffff",
                        },
                        left: {
                            size: 0,
                            color: "ffffff",
                        },
                        right: {
                            size: 0,
                            color: "ffffff",
                        },
                    }
                },
               // width: {
               //     size: 4535,
               //     type: docx.WidthType.DXA,
               // },
                rows: [
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [new docx.Paragraph({
                                    alignment: docx.AlignmentType.START,
                                    children: [
                                        new docx.TextRun("Начальник штаба игры"),
                                    ],
                                })],
                            }),
                            new docx.TableCell({
                                children: [new docx.Paragraph({
                                    alignment: docx.AlignmentType.END,
                                    children: [
                                        new docx.TextRun("____________/Корнеева Е.И",),
                                    ],
                                })],
                            }),
                        ],
                    }),
                ]
            }),
            */


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
        ]);

        //Для посредника
        if (task.check.length !== 0) {
            result2.push(
                new docx.Paragraph({
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { before: 200 },
                    children: [
                        new docx.TextRun({
                            text: 'Проверка:',
                            size: 24,
                            bold: true,
                        }),
                        new docx.TextRun({
                            text: task.check,
                            size: 24,
                            bold: true,
                        }),
                    ]
                }),
            );
        }
        if (task.notes.length !== 0) {
            result2.push(
                new docx.Paragraph({
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { before: 200 },
                    children: [
                        new docx.TextRun({
                            text: 'Примечания:',
                            size: 24,
                            bold: true,
                        }),
                        new docx.TextRun({
                            text: task.notes,
                            size: 24,
                        }),
                    ]
                }),
            );
        }

            
            
        result2.push(new docx.Paragraph("", { pageBreakAfter: true, pageBreak: true }));    
            
            
            
           // new docx.Paragraph("", { pageBreakAfter: true, pageBreak: true }), // Разрыв страницы
           // new docx.PageBreak(),


        console.log(result2)

    return result2;
}
