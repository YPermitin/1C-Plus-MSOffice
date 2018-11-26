﻿
#Область ПрограммныйИнтерфейс

Функция ЭтоПримитивныйТип(ТипЗначения) Экспорт
	
	МассивПримитивныхТипов = ПолучитьМассивПримитивныхТипов();	
	Если МассивПримитивныхТипов.Найти(ТипЗначения) = Неопределено Тогда
		Возврат Ложь;
	Иначе
		Возврат Истина;
	КОнецЕсли;
	
КонецФункции

Функция ЭтоКорректнаяДата(ЗначениеДаты) Экспорт
	
	Если ЗначениеЗаполнено(ЗначениеДаты)
		И ЗначениеДаты < Дата(1753,1,1) Тогда
		Возврат Ложь;
	Иначе
		Возврат Истина;
	КонецЕсли;
	
КонецФункции

#КонецОбласти

#Область Служебный

Функция ПолучитьМассивПримитивныхТипов()
	
	МассивПримитвныхТипов = Новый Массив;
	МассивПримитвныхТипов.Добавить(Тип("Строка"));
	МассивПримитвныхТипов.Добавить(Тип("Булево"));
	МассивПримитвныхТипов.Добавить(Тип("Дата"));
	МассивПримитвныхТипов.Добавить(Тип("Число"));
	
	Возврат МассивПримитвныхТипов;
	
КонецФункции

#КонецОбласти