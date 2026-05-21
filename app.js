// Hospital KPI Data Entry App
// Data structure from original Excel
const INDICATORS = [{"row": 2, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "نسبة الاشغال للاقسام الداخلية", "track_code": "FU-001", "ind_code": "UTZ-01", "formula": "عدد أيام إشغال الأسرة بالأقسام الداخلية خلال الفترة / (عدد الأسرة الداخلية × عدد أيام الفترة)", "num_desc": "اجمالي أيام الاشغال = مجموع عدد الايام العلاجية للمرضى = مجموع التعداد اليومي للمرضى \nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "اجمالي عدد ايام الشهر حسب الشهر (28 -30-31)", "data_source": "سجلات الدخول والخروج، نظام معلومات المستشفى.", "target": "اكثر من 85 %", "reference": "غير محدد", "responsible": "إدارة الإحصاء / السجلات الطبية", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 3, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "نسبة اشغال الرعايات المركزة", "track_code": "FU-002", "ind_code": "UTZ-02", "formula": "عدد أيام إشغال الأسرة بالرعايات المركزة خلال الفترة / (عدد أسرة الرعاية المركزة × عدد أيام الفترة)", "num_desc": "اجمالي أيام الاشغال = مجموع عدد الايام العلاجية للمرضى = مجموع التعداد اليومي للمرضى \nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "اجمالي عدد ايام الشهر حسب الشهر (28 -30-31)", "data_source": "سجلات الرعاية المركزة، نظام معلومات المستشفى.", "target": "1", "reference": "غير محدد", "responsible": "إدارة الرعاية المركزة / الإحصاء", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 4, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "نسبة اشغال حضانات حديثي الولادة", "track_code": "FU-003", "ind_code": "UTZ-03", "formula": "عدد أيام إشغال الأسرة برعاية حديثي الولادة خلال الفترة / (عدد أسرة رعاية حديثي الولادة × عدد أيام الفترة)", "num_desc": "اجمالي أيام الاشغال = مجموع عدد الايام العلاجية للمرضى = مجموع التعداد اليومي للمرضى \nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "اجمالي عدد ايام الشهر حسب الشهر (28 -30-31)", "data_source": "سجلات رعاية حديثي الولادة، نظام معلومات المستشفى.", "target": "1", "reference": "غير محدد", "responsible": "قسم حديثي الولادة / الإحصاء", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 5, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "اجمالي تردد العيادات الخارجية", "track_code": "FU-004", "ind_code": "UTZ-04", "formula": "إجمالي عدد المترددين على العيادات الخارجية خلال الفترة ", "num_desc": "لا يوجد", "den_desc": "لا يوجد", "data_source": "سجلات العيادات الخارجية، نظام معلومات المستشفى.", "target": "مقارنة بالفترات السابقة", "reference": "غير محدد", "responsible": "إدارة العيادات الخارجية / الإحصاء", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": 0, "den": 100, "result": 0, "notes": ""}, {"num": 0, "den": 100, "result": 0, "notes": 1}, {"num": 0, "den": 100, "result": 0, "notes": 1}, {"num": 0, "den": 100, "result": 0, "notes": 1}, {"num": 0, "den": 100, "result": 0, "notes": 1}, {"num": 0, "den": 100, "result": 0, "notes": 1}]}, {"row": 6, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "التردد على كل تخصص عيادة خارجية", "track_code": "FU-005", "ind_code": "UTZ-05", "formula": "عدد المترددين على تخصص معين خلال الفترة", "num_desc": "لا يوجد", "den_desc": "لا يوجد", "data_source": "سجلات العيادات الخارجية، نظام معلومات المستشفى.", "target": "مقارنة بالفترات السابقة", "reference": "غير محدد", "responsible": "إدارة العيادات الخارجية / الإحصاء", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [{"row": 7, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة :الباطنة", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 8, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : الاطفال ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 9, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : العظام", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 10, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : النساء والتوليد", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 11, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : الاسنان", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 12, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : المسالك البولية", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 13, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : الانف و الاذن و الحنجرة", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 14, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : العلاج الطبيعى  ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 15, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : المخ و الاعصاب", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 16, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : تنظيم الاسرة", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 17, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : الرمد", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 18, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : الجلدية", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 19, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : الجراحة", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 20, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : عقر حيوان ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 21, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "", "track_code": "", "ind_code": "", "formula": "تردد عيادة : رسم قلب", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}], "has_sub_items": true, "hospitals": [{"num": "", "den": "", "result": "", "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}]}, {"row": 22, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "تناسب مرضى العيادات الخارجية إلى مرضى الدخول", "track_code": "FU-006", "ind_code": "UTZ-06", "formula": "عدد مرضى العيادات الخارجية خلال الفترة / عدد المرضى الداخلين خلال نفس الفترة", "num_desc": "لا يوجد", "den_desc": "المستبعد : مرضى الاقامة أقل من 24 ساعة-مرضى الدخول مباشرة الي الاقسام الحرجة - المرضى المحولون من مستشفيات اخرى", "data_source": "سجلات العيادات والخروج، سجلات الدخول.", "target": "اكثر من 3 : 1", "reference": "غير محدد", "responsible": "إدارة الإحصاء / السجلات الطبية", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": 0, "den": "", "result": -2, "notes": ""}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}]}, {"row": 23, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "متوسط مدة الإقامة للمريض بالاقسام الداخلية", "track_code": "FU-007", "ind_code": "UTZ-07", "formula": "إجمالي عدد أيام إقامة جميع المرضى الداخلين / إجمالي عدد المرضى الخارجين من الأقسام الداخلية", "num_desc": "اجمالي عدد أيام الاقامة  = مجموع عدد الايام العلاجية للمرضى = مجموع التعداد اليومي للمرضى \nالمستبعد : الاقامة أقل من 24 ساعة- حالات الخروج حسب الطلب - مرضى الدخول للاقسام الحرجة", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات الدخول والخروج.", "target": "يحدد (مثلاً < 5 أيام)", "reference": "غير محدد", "responsible": "إدارة الإحصاء / السجلات الطبية", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": 0, "den": "", "result": -2, "notes": ""}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}]}, {"row": 24, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "متوسط مدة الإقامة للمريض باقسام الرعايات المركزة", "track_code": "FU-008", "ind_code": "UTZ-08", "formula": "إجمالي عدد أيام إقامة مرضى الرعاية المركزة / إجمالي عدد المرضى الخارجين من الرعاية المركزة", "num_desc": "اجمالي عدد أيام الاقامة  = مجموع عدد الايام العلاجية للمرضى = مجموع التعداد اليومي للمرضى \nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات الرعاية المركزة.", "target": "5-7 ايام", "reference": "غير محدد", "responsible": "إدارة الرعاية المركزة / الإحصاء", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 25, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "متوسط مدة الإقامة للمريض بقسم حديثي الولادة", "track_code": "FU-009", "ind_code": "UTZ-09", "formula": "إجمالي عدد أيام إقامة حديثي الولادة بالرعاية / عدد حديثي الولادة الخارجين من الرعاية", "num_desc": "اجمالي عدد أيام الاقامة  = مجموع عدد الايام العلاجية للمرضى = مجموع التعداد اليومي للمرضى \nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات رعاية حديثي الولادة.", "target": "5-7 ايام", "reference": "غير محدد", "responsible": "قسم حديثي الولادة / الإحصاء", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 26, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "معدل دوران الأسرة بالاقسام الداخلية", "track_code": "FU-010", "ind_code": "UTZ-10", "formula": "عدد المرضى الخارجين من الأقسام الداخلية / عدد الأسرة الداخلية", "num_desc": "المستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "لا يوجد", "data_source": "سجلات الدخول والخروج، حصر الأسرة.", "target": "يحدد لاحقاً", "reference": "غير محدد", "responsible": "إدارة الإحصاء / السجلات الطبية", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": 0, "den": "", "result": -2, "notes": ""}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}, {"num": 0, "den": "", "result": -2, "notes": -2}]}, {"row": 27, "sector": "قطاع الرعاية العلاجية", "activity": "8. كفاءة التشغيل", "name": "معدل دوران الأسرة بالاقسام بالرعاية المركزة", "track_code": "FU-011", "ind_code": "UTZ-10", "formula": "عدد المرضى الخارجين من الرعاية المركزة / عدد الأسرة بالرعاية المركزة", "num_desc": "المستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "لا يوجد", "data_source": "سجلات الدخول والخروج، حصر الأسرة.", "target": "4.5", "reference": "غير محدد", "responsible": "إدارة الإحصاء / السجلات الطبية", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 28, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "نسبة وفيات الأقسام الداخلية", "track_code": "FU-012", "ind_code": "MP-01", "formula": "(عدد حالات الوفاة بالأقسام الداخلية × 100 )/ إجمالي عدد حالات الخروج (أحياء + وفيات) من الأقسام الداخلية", "num_desc": "يشمل المرضى المحولين لقسم الرعاية وتوفوا خلال 48 ساعة من النقل \nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات الوفيات، سجلات الدخول والخروج.", "target": "أقل ما يمكن", "reference": "غير محدد", "responsible": "لجنة المراضة والوفيات / الجودة", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 29, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "نسبة الوفيات لكل قسم", "track_code": "FU-013", "ind_code": "MP-02", "formula": "عدد حالات الوفاة في كل قسم / إجمالي عدد حالات الخروج (أحياء + وفيات) من نفس القسم", "num_desc": "يشمل المرضى المحولين لقسم الرعاية وتوفوا خلال 48 ساعة من النقل\nالمستبعد : الاقامة أقل من 24 ساعة ", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات الوفيات، سجلات الدخول والخروج القسمية.", "target": "يحدد (حسب القسم)", "reference": "غير محدد", "responsible": "رؤساء الأقسام / لجنة المراضة والوفيات", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [{"row": 30, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 31, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 32, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 33, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 34, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 35, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 36, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 37, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 38, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 39, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 40, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 41, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 42, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 43, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 44, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "", "track_code": "", "ind_code": "", "formula": "القسم : ", "num_desc": "", "den_desc": "", "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}], "has_sub_items": true, "hospitals": [{"num": "", "den": "", "result": "", "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": ""}]}, {"row": 45, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "نسبة الوفيات للرعاية المركزة", "track_code": "FU-014", "ind_code": "MP-03", "formula": "عدد حالات الوفاة في الرعاية المركزة / إجمالي عدد حالات الخروج (أحياء + وفيات) من الرعاية المركزة", "num_desc": "المستبعد : المرضى المحولين من الاقسام الداخلية لقسم الرعاية وتوفوا خلال 48 ساعة من النقل -وفيات قسم الاستقبال - الاقامة أقل من 24 ساعة ", "den_desc": "عدد مرضى  الخروج شاملا الوفيات\nالمستبعد: المرضى المتوفون خلال 48 ساعة من النقل من الداخلي الي الرعاية ", "data_source": "سجلات الرعاية المركزة، سجلات الوفيات.", "target": "اقل من 3 %", "reference": "غير محدد", "responsible": "إدارة الرعاية المركزة / لجنة المراضة والوفيات", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 46, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "نسبة الوفيات للحضانات", "track_code": "FU-015", "ind_code": "MP-04", "formula": "عدد حالات الوفاة في الحضانات / إجمالي عدد حالات الخروج (أحياء + وفيات) من الحضانات", "num_desc": "المستبعد :وفيات قسم الاستقبال - الاقامة أقل من 24 ساعة ", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات الحضانات، سجلات الوفيات.", "target": "اقل من 3 %", "reference": "غير محدد", "responsible": "قسم حديثي الولادة / لجنة المراضة والوفيات", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 47, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "متوسط مدة تلقي خدمات الطوارئ (التسجيل - الخروج)", "track_code": "FU-016", "ind_code": "MP-15", "formula": "إجمالي الوقت المستغرق (من التسجيل حتى الخروج )لعينة من المرضى / عدد المرضى في العينة", "num_desc": "لا يوجد", "den_desc": "لا يوجد", "data_source": "سجلات الطوارئ، نماذج تتبع الوقت.", "target": "4 ساعات", "reference": "غير محدد", "responsible": "إدارة الطوارئ / الجودة", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 48, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "معدل حالات إعادة الدخول لقسم الطوارئ لنفس السبب خلال 72 ساعة", "track_code": "FU-017", "ind_code": "MP-12", "formula": "عدد المرضى العائدين للطوارئ لنفس السبب خلال 72 ساعة / إجمالي عدد مرضى الطوارئ", "num_desc": "المستبعد : اعادة الدخول لسبب طبي جديد / تشخيص جديد", "den_desc": "المستبعد : مرضى الطوارئ الذين دخلوا المستشفى او تم تحويلهم لمستشفى اخر", "data_source": "سجلات الطوارئ.", "target": "أقل ما يمكن", "reference": "غير محدد", "responsible": "إدارة الطوارئ / الجودة", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 49, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "معدل حالات إعادة الدخول للرعاية المركزة لنفس السبب خلال 48 ساعة", "track_code": "FU-018", "ind_code": "MP-13", "formula": "عدد الحالات العائدة للرعاية لنفس السبب خلال 48 ساعة / إجمالي عدد حالات الخروج من الرعاية", "num_desc": "المستبعد : حالات الخروج حسب الطلب - المرضى المحولين لمنشأة أخرى", "den_desc": "عدد مرضى  الخروج من الرعاية شاملا الوفيات", "data_source": "سجلات الرعاية المركزة.", "target": "أقل ما يمكن", "reference": "غير محدد", "responsible": "إدارة الرعاية المركزة / الجودة", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 50, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "معدل حالات إعادة الدخول للمستشفى لنفس السبب خلال 30 يوم", "track_code": "FU-019", "ind_code": "MP-14", "formula": "عدد حالات إعادة الدخول لنفس السبب خلال 30 يوم / إجمالي عدد حالات الخروج من المستشفى", "num_desc": "المستبعد : حالات الخروج حسب الطلب - المرضى المحولين لمنشأة أخرى", "den_desc": "عدد مرضى  الخروج شاملا الوفيات", "data_source": "سجلات الدخول والخروج.", "target": "أقل ما يمكن", "reference": "غير محدد", "responsible": "الجودة / السجلات الطبية", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 51, "sector": "قطاع الرعاية العلاجية", "activity": " الأداء الطبي", "name": "معدل حالات الخروج حسب الطلب (DAMA)", "track_code": "FU-020", "ind_code": "MP-05", "formula": "عدد حالات الخروج بناءً على طلب المريض / إجمالي عدد حالات الخروج من المستشفى", "num_desc": "المستبعد : الاقامة أقل من 24 ساعة - الوفيات", "den_desc": "المستبعد : الاقامة أقل من 24 ساعة -الوفيات", "data_source": "سجلات الخروج، نماذج الخروج حسب الطلب.", "target": "أقل 5 %", "reference": "غير محدد", "responsible": "الجودة / إدارة رضا المنتفعين", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": 0, "result": -2, "notes": ""}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}]}, {"row": 52, "sector": "الطب الوقائي", "activity": "مكافحة العدوى", "name": "معدل حدوث الالتهاب الرئوي المصاحب للتنفس الصناعي (لكل 1000 يوم تنفس)", "track_code": "FU-021", "ind_code": "IPC-03", "formula": "(إجمالي عدد حالات الالتهاب الرئوي المصاحب للتنفس الصناعي × 1000) / عدد أيام جهاز التنفس الصناعي", "num_desc": "المستبعد : العدوى المكتشفة اثناء الدخول ", "den_desc": "", "data_source": "نماذج ترصد العدوى", "target": "≤ 1 لكل 1000 يوم تنفس", "reference": "غير محدد", "responsible": "مكافحة العدوى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 53, "sector": "الطب الوقائي", "activity": "مكافحة العدوى", "name": "معدل حدوث عدوى مجرى البول المصاحب للقساطر البولية (لكل 1000 يوم قسطرة)", "track_code": "FU-022", "ind_code": "IPC-04", "formula": "(إجمالي عدد حالات عدوى مجرى البول المصاحبة للقساطر × 1000) / إجمالي عدد أيام القساطر البولية", "num_desc": "المستبعد : العدوى المكتشفة اثناء الدخول ", "den_desc": "", "data_source": "نماذج ترصد العدوى", "target": "≤ 1.5 لكل 1000 يوم قسطرة", "reference": "غير محدد", "responsible": "مكافحة العدوى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 54, "sector": "الطب الوقائي", "activity": "مكافحة العدوى", "name": "معدل عدوى موضع العمليات الجراحية", "track_code": "FU-023", "ind_code": "IPC-05", "formula": "(إجمالي عدد حالات عدوى موضع العمليات الجراحية × 100) / إجمالي عدد العمليات الجراحية التي تم متابعتها", "num_desc": "المستبعد : العدوى المكتشفة اثناء الدخول ", "den_desc": "", "data_source": "نماذج ترصد العدوى", "target": "≤ 1.2%", "reference": "غير محدد", "responsible": "مكافحة العدوى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 55, "sector": "الطب الوقائي", "activity": "مكافحة العدوى", "name": "معدل حدوث عدوى مجرى الدم المصاحب لتركيب القساطر الوريدية المركزية (لكل 1000 يوم قسطرة)", "track_code": "FU-024", "ind_code": "IPC-02", "formula": "(إجمالي عدد حالات عدوى مجرى الدم المصاحبة للقساطر × 1000) / إجمالي عدد أيام القساطر الوريدية المركزية", "num_desc": "المستبعد : العدوى المكتشفة اثناء الدخول ", "den_desc": "", "data_source": "نماذج ترصد العدوى", "target": "≤ 3 لكل 1000 يوم قسطرة", "reference": "غير محدد", "responsible": "مكافحة العدوى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 56, "sector": "قطاع تنمية المهن الطبية", "activity": "التدريب", "name": " نسبة الأطباء  المدربين على دعم الحياة القلبية التنفسية الأساسى (BLS)", "track_code": "FU-025", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "شهادات التدريب، الموارد البشرية", "target": "100 %", "reference": "القوانين واللوائح", "responsible": "إدارة التدريب", "measure_rate": "ربع سنوي", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 57, "sector": "قطاع تنمية المهن الطبية", "activity": "التدريب", "name": "نسبة التمريض المدربين على دعم الحياة القلبية التنفسية الأساسى (BLS)", "track_code": "FU-026", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "شهادات التدريب، الموارد البشرية", "target": "100 %", "reference": "القوانين واللوائح", "responsible": "إدارة التدريب", "measure_rate": "ربع سنوي", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 58, "sector": "قطاع تنمية المهن الطبية", "activity": "التدريب", "name": "اجمالى نسبة العاملين الصحيين المدربين على دعم الحياة القلبية التنفسية الأساسى (BLS)", "track_code": "FU-027", "ind_code": "WFM-05", "formula": "(عدد العاملين الصحيين المدربين على BLS × 100) / إجمالي عدد العاملين الصحيين بالمنشأة", "num_desc": "المستبعد: المشاركون  الذين لم يجتازوا التدريب بنجاح-المتدربين بجهات غير معتمدة -المتدربين الذين مضى على تدريبهم اكثر من سنتين ", "den_desc": "المستبعد: العاملين الجدد (3 شهور)", "data_source": "شهادات التدريب، الموارد البشرية", "target": "100 %", "reference": "القوانين واللوائح", "responsible": "إدارة التدريب", "measure_rate": "ربع سنوي", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 59, "sector": "قطاع تنمية المهن الطبية", "activity": "التدريب", "name": "نسبة أطباء التخصصات الحرجة المدربين على دعم الحياة القلبية المتقدم ACLS", "track_code": "FU-028", "ind_code": "WFM-08", "formula": "(عدد أطباء التخصصات الحرجة المدربين على ACLS × 100) / إجمالي عدد أطباء التخصصات الحرجة", "num_desc": "المستبعد: المشاركون  الذين لم يجتازوا التدريب بنجاح-المتدربين بجهات غير معتمدة -المتدربين الذين مضى على تدريبهم اكثر من سنتين ", "den_desc": "المستبعد: اطباء التخصصات الحرجة الجدد (3 شهور)", "data_source": "شهادات التدريب، الموارد البشرية", "target": "15 %", "reference": "غير محدد", "responsible": "إدارة التدريب", "measure_rate": "ربع سنوي", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 60, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "عدد حالات توقف القلب في كل قسم", "track_code": "FU-029", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 61, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "نسبة النجاة بعد الإنعاش القلبى  الرئوي", "track_code": "FU-030", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 62, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "نسبة وفيات حالات الإنعاش القلبى الرئوي خلال أول 24 ساعة", "track_code": "FU-031", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 63, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "إجمالي نسبة الوفيات لحالات الإنعاش القلبى الرئوي", "track_code": "FU-032", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 64, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "نسبة وفيات حالات الإنعاش القلبى الرئوي بعد 24 ساعة", "track_code": "FU-033", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 65, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "متوسط زمن الإستجابة للكود بلو (متوسط وقت الاستجابة لحالات توقف القلب)", "track_code": "FU-034", "ind_code": "MP-20", "formula": "اجمالي الوقت منذ تفعيل الكود بلو حتى وصول فريق الانعاش القلبي لمكان المريض / عدد مرات تفعيل الكود بلو", "num_desc": "لا يوجد", "den_desc": "لا يوجد", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 66, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "متوسط وقت الإنعاش القلبى الرئوي الكلى", "track_code": "FU-035", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 67, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "متوسط عدد دورات الإفاقة القلبية التنفسية لكل مريض", "track_code": "FU-036", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 68, "sector": "قطاع الرعاية العلاجية", "activity": "سلامة مرضى", "name": "نسبة الإلتزام بالتوثيق السليم للإنعاش القلبى الرئوي", "track_code": "FU-037", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 69, "sector": "الادارة المركزية لجودة الرعاية الصحية ", "activity": "متابعة المديريات", "name": "متوسط وقت الإستجابة للشكاوى", "track_code": "FU-038", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجل توثيق تفعيل اكواد الطوارئ / نموذج الكود بلو", "target": "5 دقائق", "reference": "غير محدد", "responsible": "إدارة سلامة المرضى", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 70, "sector": "الادارة المركزية لجودة الرعاية الصحية ", "activity": "متابعة المديريات", "name": "عدد شكاوى المرضى المسجلة خلال الشهر", "track_code": "FU-039", "ind_code": "PCC-01", "formula": "عدد شكاوى المرضى وأسرهم المسجلة بجميع الوسائل خلال الشهر / لا يوجد", "num_desc": "المستبعد: الشكاوى الغير موثقة والشكاوى الشفوية", "den_desc": "", "data_source": "سجلات الشكاوى", "target": "تحسن عن الشهر السابق", "reference": "غير محدد", "responsible": "إدارة رضا المنتفعين / خدمة المواطنين", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": 100, "result": -2, "notes": ""}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}, {"num": "", "den": 100, "result": -2, "notes": -2}]}, {"row": 71, "sector": "الادارة المركزية لجودة الرعاية الصحية ", "activity": "متابعة المديريات", "name": "اجمالى نسبة الشكاوى التي تم حلها", "track_code": "FU-040", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجلات الشكاوى", "target": "تحسن عن الشهر السابق", "reference": "غير محدد", "responsible": "إدارة رضا المنتفعين / خدمة المواطنين", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 72, "sector": "الادارة المركزية لجودة الرعاية الصحية ", "activity": "متابعة المديريات", "name": "نسبة الشكاوى التي تم حلها فورياً", "track_code": "FU-041", "ind_code": "", "formula": "", "num_desc": "", "den_desc": "", "data_source": "سجلات الشكاوى", "target": "تحسن عن الشهر السابق", "reference": "غير محدد", "responsible": "إدارة رضا المنتفعين / خدمة المواطنين", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}, {"row": 73, "sector": "الادارة المركزية لجودة الرعاية الصحية ", "activity": "متابعة المديريات", "name": "نسبة شكاوى المرضى التي تم حلها خلال 72 ساعة", "track_code": "FU-042", "ind_code": "PCC-02", "formula": "(عدد الشكاوى التي تم حلها خلال 72 ساعة × 100) / إجمالي عدد الشكاوى التي تم تسجيلها", "num_desc": "المستبعد: الشكاوي التي تم حلها فوريا", "den_desc": "المستبعد: الشكاوي التي تم حلها فوريا", "data_source": "سجلات الشكاوى", "target": "90 %", "reference": "غير محدد", "responsible": "إدارة رضا المنتفعين / خدمة المواطنين", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": 0, "result": -2, "notes": ""}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}, {"num": "", "den": 0, "result": -2, "notes": -2}]}, {"row": 74, "sector": "الادارة المركزية لجودة الرعاية الصحية ", "activity": "متابعة المديريات", "name": "نسبة الشكاوى التي تم حلها وعمل تغذية راجعة لمقدم الشكوى", "track_code": "FU-043", "ind_code": "PCC-03", "formula": "(عدد الشكاوى التي تم عمل تغذية راجعة لمقدم الشكوى × 100) / إجمالي الشكاوى التي تم حلها", "num_desc": "المستبعد: الشكاوي التي تم حلها فوريا", "den_desc": "المستبعد:الشكاوى التي تم حلها فوريا - الشكاوى الغير متوفر بيانات التواصل مع مقدم الشكوى", "data_source": "سجلات الشكاوى", "target": "0.75", "reference": "غير محدد", "responsible": "إدارة رضا المنتفعين / خدمة المواطنين", "measure_rate": "شهري", "report_rate": "ربع سنوي", "sub_items": [], "has_sub_items": false, "hospitals": [{"num": "", "den": "", "result": -2, "notes": ""}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}, {"num": "", "den": "", "result": -2, "notes": -2}]}];

const HOSPITALS = [
  'مستشفى 01-2026',
  'مستشفى 02-2026', 
  'مستشفى 03-2026',
  'مستشفى 04-2026',
  'مستشفى 05-2026',
  'مستشفى 06-2026'
];

// Load saved data
function loadData() {
  const saved = localStorage.getItem('hospital_kpi_data');
  if (saved) {
    return JSON.parse(saved);
  }
  return {};
}

// Save data
function saveData(data) {
  localStorage.setItem('hospital_kpi_data', JSON.stringify(data));
}

// Get current form data
function getFormData() {
  const data = loadData();
  document.querySelectorAll('.hospital-input').forEach(input => {
    const key = input.dataset.key;
    data[key] = input.value;
  });
  return data;
}

// Auto-save on change
function autoSave() {
  const data = getFormData();
  saveData(data);
  showToast('تم الحفظ تلقائياً');
}

// Show toast
function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.style.display = 'block';
  setTimeout(() => toast.style.display = 'none', 2000);
}

// Calculate result
function calculateResult(num, den) {
  const n = parseFloat(num);
  const d = parseFloat(den);
  if (!isNaN(n) && !isNaN(d) && d !== 0) {
    return (n / d).toFixed(2);
  }
  return '';
}

// Update result field
function updateResult(input) {
  const rowId = input.dataset.row;
  const hospIdx = input.dataset.hosp;
  const type = input.dataset.type; // 'num' or 'den'

  const numInput = document.querySelector(`input[data-row="${rowId}"][data-hosp="${hospIdx}"][data-type="num"]`);
  const denInput = document.querySelector(`input[data-row="${rowId}"][data-hosp="${hospIdx}"][data-type="den"]`);
  const resInput = document.querySelector(`input[data-row="${rowId}"][data-hosp="${hospIdx}"][data-type="result"]`);

  if (numInput && denInput && resInput) {
    const result = calculateResult(numInput.value, denInput.value);
    resInput.value = result;
    autoSave();
  }
}

// Build the form
function buildForm() {
  const container = document.getElementById('indicators-container');
  const savedData = loadData();

  let currentSector = '';
  let currentActivity = '';
  let sectorDiv = null;
  let activityDiv = null;

  INDICATORS.forEach((ind, idx) => {
    // Sector grouping
    if (ind.sector !== currentSector) {
      currentSector = ind.sector;
      sectorDiv = document.createElement('div');
      sectorDiv.className = 'sector-section';
      sectorDiv.innerHTML = `<h2 class="sector-title">${currentSector}</h2>`;
      container.appendChild(sectorDiv);
      currentActivity = '';
    }

    // Activity grouping
    if (ind.activity !== currentActivity) {
      currentActivity = ind.activity;
      activityDiv = document.createElement('div');
      activityDiv.className = 'activity-section';
      activityDiv.innerHTML = `<h3 class="activity-title">${currentActivity}</h3>`;
      sectorDiv.appendChild(activityDiv);
    }

    // Indicator card
    const card = document.createElement('div');
    card.className = 'indicator-card';
    card.id = `indicator-${idx}`;

    let html = `
      <div class="indicator-header" onclick="toggleCard(${idx})">
        <div class="indicator-title">
          <span class="toggle-icon">▼</span>
          <strong>${ind.name || ind.formula}</strong>
          ${ind.track_code ? `<span class="badge">${ind.track_code}</span>` : ''}
          ${ind.ind_code ? `<span class="badge badge-secondary">${ind.ind_code}</span>` : ''}
        </div>
        <div class="indicator-meta">
          ${ind.formula ? `<small>${ind.formula}</small>` : ''}
        </div>
      </div>
      <div class="indicator-body" id="body-${idx}">
        <div class="indicator-info">
          ${ind.num_desc ? `<div class="info-row"><label>البسط:</label> <span>${ind.num_desc}</span></div>` : ''}
          ${ind.den_desc ? `<div class="info-row"><label>المقام:</label> <span>${ind.den_desc}</span></div>` : ''}
          ${ind.data_source ? `<div class="info-row"><label>مصادر البيانات:</label> <span>${ind.data_source}</span></div>` : ''}
          ${ind.target ? `<div class="info-row"><label>المستهدف:</label> <span>${ind.target}</span></div>` : ''}
          ${ind.responsible ? `<div class="info-row"><label>المسئول:</label> <span>${ind.responsible}</span></div>` : ''}
        </div>
        <div class="hospitals-grid">
    `;

    HOSPITALS.forEach((hosp, hIdx) => {
      const hospData = ind.hospitals[hIdx] || {};
      const baseKey = `${idx}_${hIdx}`;

      html += `
        <div class="hospital-box">
          <div class="hospital-name">${hosp}</div>
          <div class="input-group">
            <label>قيمة البسط</label>
            <input type="number" class="hospital-input" data-key="${baseKey}_num" data-row="${idx}" data-hosp="${hIdx}" data-type="num" 
              value="${savedData[baseKey + '_num'] || hospData.num || ''}" 
              onchange="updateResult(this)" placeholder="0">
          </div>
          <div class="input-group">
            <label>قيمة المقام</label>
            <input type="number" class="hospital-input" data-key="${baseKey}_den" data-row="${idx}" data-hosp="${hIdx}" data-type="den" 
              value="${savedData[baseKey + '_den'] || hospData.den || ''}" 
              onchange="updateResult(this)" placeholder="0">
          </div>
          <div class="input-group">
            <label>ناتج المؤشر</label>
            <input type="text" class="hospital-input result-input" data-key="${baseKey}_result" data-row="${idx}" data-hosp="${hIdx}" data-type="result" 
              value="${savedData[baseKey + '_result'] || hospData.result || ''}" 
              onchange="autoSave()" placeholder="0.00" readonly>
          </div>
          <div class="input-group">
            <label>الملاحظات</label>
            <input type="text" class="hospital-input" data-key="${baseKey}_notes" data-row="${idx}" data-hosp="${hIdx}" data-type="notes" 
              value="${savedData[baseKey + '_notes'] || hospData.notes || ''}" 
              onchange="autoSave()" placeholder="ملاحظات...">
          </div>
        </div>
      `;
    });

    html += `</div>`; // close hospitals-grid

    // Sub items
    if (ind.has_sub_items && ind.sub_items.length > 0) {
      html += `<div class="sub-items">`;
      ind.sub_items.forEach((sub, sIdx) => {
        html += `
          <div class="sub-item">
            <div class="sub-formula">${sub.formula}</div>
            <div class="hospitals-grid sub-hospitals">
        `;
        HOSPITALS.forEach((hosp, hIdx) => {
          const subHospData = sub.hospitals[hIdx] || {};
          const subKey = `${idx}_sub${sIdx}_${hIdx}`;
          html += `
            <div class="hospital-box sub-box">
              <div class="hospital-name">${hosp}</div>
              <div class="input-group">
                <label>قيمة البسط</label>
                <input type="number" class="hospital-input" data-key="${subKey}_num" data-row="${idx}_sub${sIdx}" data-hosp="${hIdx}" data-type="num" 
                  value="${savedData[subKey + '_num'] || subHospData.num || ''}" 
                  onchange="updateResult(this)" placeholder="0">
              </div>
              <div class="input-group">
                <label>قيمة المقام</label>
                <input type="number" class="hospital-input" data-key="${subKey}_den" data-row="${idx}_sub${sIdx}" data-hosp="${hIdx}" data-type="den" 
                  value="${savedData[subKey + '_den'] || subHospData.den || ''}" 
                  onchange="updateResult(this)" placeholder="0">
              </div>
              <div class="input-group">
                <label>ناتج المؤشر</label>
                <input type="text" class="hospital-input result-input" data-key="${subKey}_result" data-row="${idx}_sub${sIdx}" data-hosp="${hIdx}" data-type="result" 
                  value="${savedData[subKey + '_result'] || subHospData.result || ''}" 
                  onchange="autoSave()" placeholder="0.00" readonly>
              </div>
              <div class="input-group">
                <label>الملاحظات</label>
                <input type="text" class="hospital-input" data-key="${subKey}_notes" data-row="${idx}_sub${sIdx}" data-hosp="${hIdx}" data-type="notes" 
                  value="${savedData[subKey + '_notes'] || subHospData.notes || ''}" 
                  onchange="autoSave()" placeholder="ملاحظات...">
              </div>
            </div>
          `;
        });
        html += `</div></div>`;
      });
      html += `</div>`;
    }

    html += `</div>`; // close indicator-body
    card.innerHTML = html;
    activityDiv.appendChild(card);
  });
}

function toggleCard(idx) {
  const body = document.getElementById(`body-${idx}`);
  const icon = document.querySelector(`#indicator-${idx} .toggle-icon`);
  if (body.style.display === 'none') {
    body.style.display = 'block';
    icon.textContent = '▼';
  } else {
    body.style.display = 'none';
    icon.textContent = '▶';
  }
}

// Export to Excel
function exportExcel() {
  const data = getFormData();

  // Prepare workbook structure matching original
  const wb = XLSX.utils.book_new();
  wb.Props = {
    Title: "نموذج إدخال بيانات مقارنة مؤشرات مستشفيات",
    Subject: "مؤشرات المستشفيات",
    Author: "Hospital KPI System",
    CreatedDate: new Date()
  };

  // Build worksheet data array
  const wsData = [];

  // Row 0: Main headers
  const row0 = new Array(38).fill('');
  row0[2] = 'اسم المؤشر';
  row0[3] = 'كود المؤشر بإدارة المتابعة';
  row0[4] = 'كود المؤشر';
  row0[5] = 'المعادلة (الوصف)';
  row0[6] = 'تعليمات تجميع البيانات';
  row0[8] = 'مستشفى 01-2026';
  row0[11] = 'الملاحظات';
  row0[12] = 'مستشفى 02-2026';
  row0[15] = 'الملاحظات';
  row0[16] = 'مستشفى 03-2026';
  row0[19] = 'الملاحظات';
  row0[20] = 'مستشفى 04-2026';
  row0[23] = 'الملاحظات';
  row0[24] = 'مستشفى 05-2026';
  row0[27] = 'الملاحظات';
  row0[28] = 'مستشفى 06-2026';
  row0[31] = 'الملاحظات';
  row0[32] = 'مصادر البيانات';
  row0[33] = 'المستهدف';
  row0[34] = 'المرجع';
  row0[35] = 'المسئول';
  row0[36] = 'معدل القياس';
  row0[37] = 'معدل الإبلاغ';
  wsData.push(row0);

  // Row 1: Sub headers
  const row1 = new Array(38).fill('');
  row1[0] = 'القطاع';
  row1[1] = 'النشاط';
  row1[6] = 'البسط';
  row1[7] = 'المقام';
  row1[8] = 'قيمة البسط';
  row1[9] = 'قيمة المقام';
  row1[10] = 'ناتج المؤشر';
  row1[12] = 'قيمة البسط';
  row1[13] = 'قيمة المقام';
  row1[14] = 'ناتج المؤشر';
  row1[16] = 'قيمة البسط';
  row1[17] = 'قيمة المقام';
  row1[18] = 'ناتج المؤشر';
  row1[20] = 'قيمة البسط';
  row1[21] = 'قيمة المقام';
  row1[22] = 'ناتج المؤشر';
  row1[24] = 'قيمة البسط';
  row1[25] = 'قيمة المقام';
  row1[26] = 'ناتج المؤشر';
  row1[28] = 'قيمة البسط';
  row1[29] = 'قيمة المقام';
  row1[30] = 'ناتج المؤشر';
  wsData.push(row1);

  // Data rows
  INDICATORS.forEach((ind, idx) => {
    const row = new Array(38).fill('');
    row[0] = ind.sector;
    row[1] = ind.activity;
    row[2] = ind.name;
    row[3] = ind.track_code;
    row[4] = ind.ind_code;
    row[5] = ind.formula;
    row[6] = ind.num_desc;
    row[7] = ind.den_desc;
    row[32] = ind.data_source;
    row[33] = ind.target;
    row[34] = ind.reference;
    row[35] = ind.responsible;
    row[36] = ind.measure_rate;
    row[37] = ind.report_rate;

    // Hospital data
    const hMap = [{c:[8,9,10,11]}, {c:[12,13,14,15]}, {c:[16,17,18,19]}, {c:[20,21,22,23]}, {c:[24,25,26,27]}, {c:[28,29,30,31]}];
    hMap.forEach((hm, hIdx) => {
      const baseKey = `${idx}_${hIdx}`;
      row[hm.c[0]] = data[baseKey + '_num'] || ind.hospitals[hIdx].num || '';
      row[hm.c[1]] = data[baseKey + '_den'] || ind.hospitals[hIdx].den || '';
      row[hm.c[2]] = data[baseKey + '_result'] || ind.hospitals[hIdx].result || '';
      row[hm.c[3]] = data[baseKey + '_notes'] || ind.hospitals[hIdx].notes || '';
    });

    wsData.push(row);

    // Sub items
    if (ind.has_sub_items) {
      ind.sub_items.forEach((sub, sIdx) => {
        const subRow = new Array(38).fill('');
        subRow[0] = sub.sector;
        subRow[1] = sub.activity;
        subRow[5] = sub.formula;
        subRow[6] = sub.num_desc;
        subRow[7] = sub.den_desc;

        hMap.forEach((hm, hIdx) => {
          const subKey = `${idx}_sub${sIdx}_${hIdx}`;
          subRow[hm.c[0]] = data[subKey + '_num'] || sub.hospitals[hIdx].num || '';
          subRow[hm.c[1]] = data[subKey + '_den'] || sub.hospitals[hIdx].den || '';
          subRow[hm.c[2]] = data[subKey + '_result'] || sub.hospitals[hIdx].result || '';
          subRow[hm.c[3]] = data[subKey + '_notes'] || sub.hospitals[hIdx].notes || '';
        });

        wsData.push(subRow);
      });
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Set column widths
  ws['!cols'] = [
    {wch:25}, {wch:20}, {wch:35}, {wch:12}, {wch:10}, {wch:50}, {wch:40}, {wch:40},
    {wch:12}, {wch:12}, {wch:12}, {wch:15},
    {wch:12}, {wch:12}, {wch:12}, {wch:15},
    {wch:12}, {wch:12}, {wch:12}, {wch:15},
    {wch:12}, {wch:12}, {wch:12}, {wch:15},
    {wch:12}, {wch:12}, {wch:12}, {wch:15},
    {wch:12}, {wch:12}, {wch:12}, {wch:15},
    {wch:30}, {wch:15}, {wch:12}, {wch:25}, {wch:10}, {wch:12}
  ];

  // Add merges for main headers (row 0)
  ws['!merges'] = [
    // Hospital 01
    {s:{r:0,c:8}, e:{r:0,c:10}},
    // Hospital 02
    {s:{r:0,c:12}, e:{r:0,c:14}},
    // Hospital 03
    {s:{r:0,c:16}, e:{r:0,c:18}},
    // Hospital 04
    {s:{r:0,c:20}, e:{r:0,c:22}},
    // Hospital 05
    {s:{r:0,c:24}, e:{r:0,c:26}},
    // Hospital 06
    {s:{r:0,c:28}, e:{r:0,c:30}},
    // Notes columns are single
    // Instructions span
    {s:{r:0,c:6}, e:{r:0,c:7}}
  ];

  // RTL
  ws['!dir'] = 'rtl';

  XLSX.utils.book_append_sheet(wb, ws, 'مستشفى ');

  const today = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `نموذج_مؤشرات_المستشفيات_${today}.xlsx`);
  showToast('تم تصدير الملف بنجاح');
}

// Clear all data
function clearAll() {
  if (confirm('هل أنت متأكد من حذف جميع البيانات؟')) {
    localStorage.removeItem('hospital_kpi_data');
    location.reload();
  }
}

// Search
function searchIndicators() {
  const query = document.getElementById('search').value.toLowerCase();
  document.querySelectorAll('.indicator-card').forEach(card => {
    const text = card.textContent.toLowerCase();
    if (text.includes(query)) {
      card.style.display = 'block';
      // Expand if searching
      if (query) {
        const body = card.querySelector('.indicator-body');
        const icon = card.querySelector('.toggle-icon');
        body.style.display = 'block';
        icon.textContent = '▼';
      }
    } else {
      card.style.display = 'none';
    }
  });
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  buildForm();

  // Register service worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('service-worker.js')
      .then(reg => console.log('SW registered'))
      .catch(err => console.log('SW error', err));
  }
});
