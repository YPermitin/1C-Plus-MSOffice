## Работа с Access

# Подготовка

Для работы с базами Access необходимо на сервере 1С (в клиент-серверном режиме) или на компьютере, где запускается клиентское приложение (файловая база, толстый клиент), установить один из следующих компонентов:
- [Полный пакет Microsoft Access](https://infostart.ru/redirect.php?url=aHR0cHM6Ly9wcm9kdWN0cy5vZmZpY2UuY29tL3J1LXJ1L2NvbXBhcmUtYWxsLW1pY3Jvc29mdC1vZmZpY2UtcHJvZHVjdHM/dGFiPTEmYW1wO09DSUQ9QUlENjc5NDcxX09PX0RMQ19RNHJlZnJlc2g=).
- [Microsoft Access Database Engine 2016 Redistributable](https://infostart.ru/redirect.php?url=aHR0cHM6Ly93d3cubWljcm9zb2Z0LmNvbS9lbi11cy9kb3dubG9hZC9kZXRhaWxzLmFzcHg/aWQ9NTQ5MjA=).
- [Microsoft Access 2016 Runtime](https://infostart.ru/redirect.php?url=aHR0cHM6Ly93d3cubWljcm9zb2Z0LmNvbS9ydS1SVS9kb3dubG9hZC9kZXRhaWxzLmFzcHg/aWQ9NTAwNDA=).

Установка Microsoft Access 2016 Runtime не требует покупки дополнительных лицензий, т.к. этот пакет содержит лишь среду выполнения, которая используется для запуска уже готовых решений. Средства разработки в ней отсутствуют. При этом в состав пакета также входит установщик ODBC-драйвера, который нам и нужен.

Подробнее про лицензировние можно [прочитать здесь](https://infostart.ru/redirect.php?url=aHR0cHM6Ly93d3cucmlwdGlkZWhvc3RpbmcuY29tL2Jsb2cvdGFnL21pY3Jvc29mdC1saWNlbnNpbmcv), в том числе и в контексте MS Access.

# Простые примеры

Чтение данных из базы Access из кода встроенного языка 1С:Предприятия можно выполнять следующим образом.
```bsl
ФайлБазы = "<путь к базе *.accdb/*.mdb>";

// Инициализация подключения к базе
СтрокаПодключения = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + ФайлБазы;
Connection = Новый COMОбъект("ADODB.Connection");	
Connection.Open(СтрокаПодключения);

// Формируем команды чтения данных из таблицы "Выгрузка_результата_компоновки"
Command = Новый COMОбъект("ADODB.Command");
Command.ActiveConnection = Connection;
Command.CommandText = "Select * FROM Выгрузка_результата_компоновки";
Command.CommandType = 1;
RecordSet = Новый COMОбъект("ADODB.RecordSet");
RecordSet = Command.Execute();

// Считываем все поля и выводим пользователю
ЗначенияСтрокой ="";
Пока RecordSet.EOF() = 0 Цикл
	
	Для НомерПоля = 0 по  Recordset.Fields.Count - 1 цикл
		ЗначенияСтрокой = ЗначенияСтрокой + " " + Recordset.Fields(НомерПоля).Value;
	КонецЦикла;
	
	Сообщить(ЗначенияСтрокой);		
	ЗначенияСтрокой = "";
	
	RecordSet.MoveNext();
	
КонецЦикла;

// Освобождаем ресурсы
RecordSet.Close();
Connection.Close();
```

Для выгрузки пример будет походим, но файл базы данных должен быть уже готовым к операции, в т.ч. и с необходимыми таблицами.
```bsl
Запрос = Новый Запрос;
Запрос.Текст = 
	"ВЫБРАТЬ
	|	&ТекущаяДата КАК Дата,
	|	""Привет из Access"" КАК Строка";	
Запрос.УстановитьПараметр("ТекущаяДата", ТекущаяДата());
ТаблицаИсточник = Запрос.Выполнить().Выгрузить();

// Подготовленная база Access для выгрузки
ПутьКБазе = "C:\Access\ПростаяВыгрузка.accdb";

СтрокаПодключения = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + ПутьКБазе;
СоединениеКБазе = Новый COMОбъект("ADODB.Connection");	
СоединениеКБазе.Open(СтрокаПодключения);

ИмяТаблицы = "ПростаяВыгрузка";
Запись = Новый COMОбъект("ADODB.RecordSet");	
ТекстЗапроса = "SELECT * FROM " + ИмяТаблицы;	
Запись.Open(
	// Текст запроса 
	ТекстЗапроса, 
	// Соединение с базой
	СоединениеКБазе,
	// Указывает тип курсора, используемого в записей объекта.
	// CursorType (https://docs.microsoft.com/ru-ru/sql/ado/reference/ado-api/cursortypeenum?view=sql-server-2017)
	// 1 = adOpenKeyset. Использует курсор набора ключей. 
	1, 			  
	// Тип блокировки
	// LockTypeEnum (https://docs.microsoft.com/ru-ru/sql/ado/reference/ado-api/open-method-ado-recordset?view=sql-server-2017)
	// 3 = adLockOptimistic (Указывает, оптимистической блокировки, записей.)
	3
);

// Добавляем записи в таблицу базы Access
//	В исходном файле первая колонка содержит дату,
//	а во второй сохраняем строку.
Для Каждого СтрокаТаблицы Из ТаблицаИсточник Цикл
	
	Запись.AddNew();
	Запись.Fields(0).Value = СтрокаТаблицы.Дата;
	Запись.Fields(1).Value = СтрокаТаблицы.Строка;
	Запись.UpDate();
	
КонецЦикла;	

СоединениеКБазе.Close();
СоединениеКБазе = Неопределено;
```

# Подсистема работы с Access

Описание в работе...
