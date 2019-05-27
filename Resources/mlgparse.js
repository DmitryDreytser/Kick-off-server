function logoptionsselect()
{
var logdata = [
['Refs','','Все'],
['Refs','RefOpen','Открыт'],
['Refs','RefWrite','Записан'],
['Refs','RefNew','Создан'],
['Refs','RefMarkDel','Помечен на удаление'],
['Refs','RefUnmarkDel','Снята пометка удаления'],
['Refs','RefDel','Удален'],
['Refs','RefGrpMove','Перенесен в другую группу'],
['Refs','RefAttrWrite','Значение реквизита записано'],
['Refs','RefAttrDel','Значение реквизита удалено'],

['Docs','','Все'],
['Docs','DocNew','Cоздан'],
['Docs','DocOpen','Открыт'],
['Docs','DocWrite','Записан'],
['Docs','DocWriteNew','Записан новый'],
['Docs','DocNotWrite','Не записан'],
['Docs','DocPassed','Проведен'],
['Docs','DocBackPassed','Проведен задним числом'],
['Docs','DocNotPassed','Не проведен'],
['Docs','DocMakeNotPassed','Сделан не проведенным'],
['Docs','DocWriteAndPassed','Записан и проведен'],
['Docs','DocWriteAndRepassed','Записан и проведен задним числом'],
['Docs','DocTimeChanged','Изменено время'],
['Docs','DocOperOn','Проводки включены'],
['Docs','DocOperOff','Проводки выключены'],
['Docs','DocMarkDel','Помечен на удаление'],
['Docs','DocUnmarkDel','Снята пометка удаления'],
['Docs','DocDel','Удален'],

['Distr','','Все'],
['Distr','DistBatchErr','Ошибка автообмена в пакетном режиме'],
['Distr','DistDnldBeg','Начата выгрузка изменений данных'],
['Distr','DistDnldSuc','Выгрузка изменений данных успешно завершена'],
['Distr','DistDnldFail','Выгрузка изменений данных не выполнена'],
['Distr','DistDnlErr','Ошибка выгрузки изменений данных'],
['Distr','DistUplBeg','Начата загрузка изменений данных'],
['Distr','DistUplSuc','Загрузка изменений данных успешно завершена'],
['Distr','DistUplFail','Загрузка изменений данных не выполнена'],
['Distr','DistUplErr','Ошибка загрузки изменений данных'],
['Distr','DistUplStatus','Загрузка изменений данных'],
['Distr','DistDnldPrimBeg','Первичная выгрузка периферийной ИБ'],
['Distr','DistDnldPrimSuc','Первичная выгрузка периферийной ИБ успешно завершена'],
['Distr','DistInit','Распределенная ИБ инициализирована'],
['Distr','DistPIBCreat','Создана периферийная ИБ'],
['Distr', 'DistAEParam', 'Изменены параметры автообмена'],

['Restruct','', 'Все'],
['Restruct', 'RestructSaveMD', 'Запись измененной конфигурации'],
['Restruct', 'RestructStart', 'Начало реструктуризации'],
['Restruct', 'RestructCopy', 'Начато копирование результатов реструктуризации'],
['Restruct', 'RestructAcptEnd', 'Реструктуризация завершена'],
['Restruct', 'RestructStatus', 'Статус реструктуризации'],
['Restruct', 'RestructAnalys', 'Анализ информации'],
['Restruct', 'RestructStartWarn', 'Предупреждение'],
['Restruct', 'RestructErr', 'Ошибка при реструктуризации'],
['Restruct', 'RestructCritErr', 'Критическая ошибка при реструктуризации'],

['Grbgs', '', 'Все'],
['Grbgs', 'GrbgSyntaxErr', 'Синтаксическая ошибка'],
['Grbgs', 'GrbgNewPerBuhTot', 'Бухгалтерские итоги рассчитаны'],
['Grbgs', 'GrbgRclcAllBuhTot', 'Полный пересчет бухгалтерских итогов'],
['Grbgs', 'GrbgRuntimeErr', 'Ошибка времени выполнения']

] ;

var TypeOperation = document.getElementById('objecttype').value;
var enttypelist = document.dbaction.eventtype;
var count_option = document.dbaction.eventtype.length;
	for ( i = 0; i < count_option; i++ ) 
	{
		enttypelist.remove(enttypelist.options[i]);
	}
	
	for ( i = 0; i < logdata.length; i++ ) 
		if (logdata[i][0] == TypeOperation)
				enttypelist[enttypelist.length] = new Option(logdata[i][2],logdata[i][1],false,(enttypelist.length == 0));	
	return false;
}