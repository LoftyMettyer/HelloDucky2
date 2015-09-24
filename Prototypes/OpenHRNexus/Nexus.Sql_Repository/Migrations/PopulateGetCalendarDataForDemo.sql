
ALTER VIEW GetCalendarData
AS
SELECT [id]
, column20 AS [title]
, convert(bit, 0) AS [allDay]
, column17 AS [start]
, column18 AS [end]
, 'www.urllinkgoeshere.com' AS [url]
, ISNULL([column28], '') AS [className]
, convert(bit, 0) AS [editable]
, convert(bit, 0) AS [startEditable]
, convert(bit, 0) AS [durationEditable]
, convert(bit, 0) AS [overlap]
, NULL AS [constraint]
, ISNULL([column27], '') AS [color]
, '' AS [backgroundcolor]
, '' AS [bordercolor]
, '' AS [textcolor]
FROM UserDefined2

INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-06','2015-07-06','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-06','2015-07-06','AGM', 'Purple', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-07','2015-07-10','.NET Best Practices and Design Patterns - 4 Days ', 'Red', 'Work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-07','2015-07-07','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-08','2015-07-08','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-09','2015-07-09','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-09','2015-07-09','MWM; Venue Rotation', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-10','2015-07-10','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-10','2015-07-12','Camping; Gilwell Park', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-10','2015-07-10','Uni end of term', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-11','2015-07-11','Dinner Out', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-13','2015-07-13','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-14','2015-07-14','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-14','2015-07-14','Working From Home', 'blue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-14','2015-07-17','Car in for repair', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-15','2015-07-15','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-16','2015-07-16','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-16','2015-07-16','Dentist', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-16','2015-07-16','Cinema; The odyssey', 'grey', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-17','2015-07-17','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-17','2015-07-21','Holiday', 'Darkblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-17','2015-07-17','Flight Out; London Luton Airport', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-20','2015-07-20','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-20','2015-07-20','Visual Studio 2015 Final Release Event', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-20','2015-07-31','Harry Off', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-21','2015-07-21','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-21','2015-07-21','FlightBack; Aeroport de Palma de Mallorca (PMI) (Aeroport de Palma Palma, Balearische Inseln España)', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-22','2015-07-22','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-23','2015-07-23','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-23','2015-07-23','Classics on the Common', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-23','2015-07-23','MWM; Venue Rotation', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-24','2015-07-24','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-25','2015-07-25','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-26','2015-07-26','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-26','2015-07-26','Arsenal vs Wolfsburg; Emirates Stadium', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-27','2015-07-27','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-27','2015-08-07','Roberto Off', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-27','2015-07-27','Holiday PM', 'Darkblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-27','2015-07-27','Orthodontist visit', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-28','2015-07-28','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-28','2015-07-28','Drive wifey to her do', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-29','2015-07-29','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-29','2015-07-29','Using Node.js with Visual Studio Code Jump Start', 'green', 'Personal')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-30','2015-07-30','Daily Scrum', 'lightblue', 'work')
INSERT UserDefined2_static ([column17], [column18], [column20], [column27], [column28]) VALUES ( '2015-07-31','2015-07-31','Daily Scrum', 'lightblue', 'work')
