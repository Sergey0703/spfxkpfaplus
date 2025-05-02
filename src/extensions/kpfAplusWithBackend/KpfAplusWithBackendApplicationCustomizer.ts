// src/extensions/kpfAplusWithBackend/KpfAplusWithBackendApplicationCustomizer.ts

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

// Импорты PnP JS
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

// Импорт логгера
import { Logger } from './services/Logger';

import styles from './AppCustomizer.module.scss';
import * as strings from 'KpfAplusWithBackendApplicationCustomizerStrings';

const LOG_SOURCE: string = 'KpfAplusWithBackendApplicationCustomizer';

/**
 * Интерфейс свойств приложения
 */
export interface IKpfAplusWithBackendApplicationCustomizerProperties {
  // Сообщение, которое может быть настроено
  testMessage: string;
  // Название списка для обработки
  listName?: string;
  // Целевое поле для фильтрации
  targetField?: string;
  // Целевое значение для фильтрации
  targetValue?: string;
}

/**
 * Интерфейс для запросов к API
 */
export interface IApiRequestData {
  listName?: string;
  [key: string]: unknown;
}

/**
 * Интерфейс для результатов обработки элементов
 */
export interface IProcessedItem {
  id: number;
  title: string;
  processed: boolean;
  timestamp: string;
}

/**
 * Интерфейсы для HTTP запросов (так как они отсутствуют в вашей версии SPFx)
 */
export interface IHttpRequest {
  url: string;
  method: string;
  headers?: Record<string, string>;
  body?: string;
}

export interface IHttpResponse {
  status: number;
  headers: Record<string, string>;
  body: string;
}

/** 
 * Application Customizer с REST API и пользовательским интерфейсом 
 */
export default class KpfAplusWithBackendApplicationCustomizer
  extends BaseApplicationCustomizer<IKpfAplusWithBackendApplicationCustomizerProperties> {

  // Плейсхолдер для отображения UI
  private _topPlaceholder: PlaceholderContent | undefined;
  
  // PnP SP объект для работы с SharePoint
  private sp: SPFI;
  
  // Название списка-триггера для запуска операций
  private readonly TRIGGER_LIST_NAME = "APlusTrigger";
  
  // Интервал таймера - используем number вместо NodeJS.Timeout
  private timerInterval: number | null = null;
  
  // По умолчанию таймер отключен
  private isTimerActive: boolean = false;
  
  // Локальный ключ для хранения информации о последнем запуске
  private readonly LOCAL_STORAGE_KEY = 'KPF_DailyService_LastRunDate';
  
  // Название списка настроек
  private readonly SETTINGS_LIST_NAME = "KPFAPlus_Settings";

  /**
   * Карта соответствия имен списков
   * Используется для маппинга имен списков из запроса в реальные имена списков SharePoint
   */
  private readonly LIST_NAME_MAPPING: { [key: string]: string } = {
    'Timetable': 'TestTasks',
    'Schedule': 'TestTasks',
    'Расписание': 'TestTasks',
    'Tasks': 'TestTasks',      // Добавляем маппинг для "Tasks"
    'TestTasks': 'TestTasks'   // Добавляем маппинг для самого себя
  };

  @override
  public async onInit(): Promise<void> {
    try {
      // Инициализируем логгер
      Logger.initialize(this.context);
      
      // Инициализируем PnP SP
      this.sp = spfi().using(SPFx(this.context));
      
      // Загружаем настройки таймера из SharePoint списка
      await this.loadTimerSettings();
      
      // Логируем начало инициализации
      await Logger.info(LOG_SOURCE, `Инициализация расширения KPF A-Plus`, {
        version: '1.0.0',
        userAgent: window.navigator.userAgent,
        user: this.context.pageContext.user.displayName,
        url: window.location.href,
        timerActive: this.isTimerActive
      });
      
      // Стандартное логирование (для совместимости)
      Log.info(LOG_SOURCE, `Инициализация ${strings.Title}`);

      // Установка обработчиков для плейсхолдеров
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

      // Первая отрисовка UI
      this._renderPlaceHolders();

      // Запускаем таймер для проверки триггер-списка, только если он активен
      if (this.isTimerActive) {
        this.startTimer();
      } else {
        console.log("Таймер отключен согласно настройкам в SharePoint");
      }
      
      // Логируем успешную инициализацию
      await Logger.info(LOG_SOURCE, 'Расширение успешно инициализировано');

      return Promise.resolve();
    } catch (error) {
      // Логируем ошибку инициализации
      await Logger.critical(LOG_SOURCE, `Ошибка при инициализации расширения: ${error}`, {
        stack: error instanceof Error ? error.stack : undefined,
        properties: this.properties
      });
      throw error;
    }
  }

  /**
   * Новый метод для обработки HTTP запросов к API
   * @param request HTTP запрос
   * @returns HTTP ответ
   */
  @override
  public async onHttpRequest(request: IHttpRequest): Promise<IHttpResponse> {
    try {
      // Получаем URL и метод запроса
      const url = request.url.toLowerCase();
      const method = request.method.toUpperCase();
      
      // Логируем получение HTTP запроса
      await Logger.info(LOG_SOURCE, `Получен HTTP запрос`, { 
        url,
        method,
        timestamp: new Date().toISOString()
      });

      // Проверяем, является ли запрос вызовом API processlist
      if (url.indexOf('/kpfaplus/api/processlist') === url.length - '/kpfaplus/api/processlist'.length && method === 'POST') {
        // Парсим тело запроса
        let requestData: IApiRequestData = {};
        try {
          if (request.body) {
            requestData = JSON.parse(request.body);
            await Logger.info(LOG_SOURCE, `Параметры запроса API`, {
              ...requestData,
              timestamp: new Date().toISOString()
            });
          }
        } catch (e) {
          await Logger.error(LOG_SOURCE, `Ошибка парсинга JSON: ${e}`, { 
            body: request.body 
          });
          
          // Возвращаем ошибку парсинга
          return {
            status: 400,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              success: false,
              error: `Ошибка парсинга JSON: ${e instanceof Error ? e.message : String(e)}`
            })
          };
        }
        
        // Получаем имя списка из запроса или используем TestTasks по умолчанию
        const requestedListName = requestData.listName || 'TestTasks';
        
        // Получаем реальное имя списка с использованием маппинга
        const actualListName = this.getActualListName(requestedListName.toString());
        
        // Выполняем бизнес-логику с реальным именем списка
        try {
          const result = await this.processListBusinessLogic(actualListName, requestData);
          
          // Логируем успешное выполнение API запроса
          await Logger.info(LOG_SOURCE, `API запрос успешно выполнен`, { 
            requestedListName,
            actualListName,
            processedItems: result.length,
            timestamp: new Date().toISOString()
          });
          
          // Возвращаем успешный результат
          return {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              success: true,
              requestedList: requestedListName,
              actualList: actualListName,
              processedItems: result.length,
              items: result,
              timestamp: new Date().toISOString()
            })
          };
        } catch (processError) {
          // Логируем ошибку обработки
          await Logger.error(LOG_SOURCE, `Ошибка при выполнении бизнес-логики: ${processError}`, {
            requestedListName,
            actualListName,
            error: processError instanceof Error ? processError.message : String(processError),
            stack: processError instanceof Error ? processError.stack : undefined,
            timestamp: new Date().toISOString()
          });
          
          // Возвращаем ошибку обработки
          return {
            status: 500,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              success: false,
              error: processError instanceof Error ? processError.message : String(processError)
            })
          };
        }
      }
      
      // Если запрос не соответствует ни одному из API-методов
      return {
        status: 404,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          success: false,
          error: 'API метод не найден'
        })
      };
    } catch (error) {
      // Логируем неожиданную ошибку
      await Logger.critical(LOG_SOURCE, `Неожиданная ошибка при обработке HTTP запроса: ${error}`, {
        url: request.url,
        method: request.method,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        timestamp: new Date().toISOString()
      });
      
      // Возвращаем ошибку сервера
      return {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          success: false,
          error: `Внутренняя ошибка сервера: ${error instanceof Error ? error.message : String(error)}`
        })
      };
    }
  }

  /**
   * Загружает настройки таймера из списка SharePoint
   */
  private async loadTimerSettings(): Promise<void> {
    try {
      // Проверяем существование списка настроек
      try {
        // Получаем настройки
        const settings = await this.sp.web.lists.getByTitle(this.SETTINGS_LIST_NAME).items
          .filter("Title eq 'TimerSettings'")();
        
        if (settings.length > 0) {
          const setting = settings[0];
          
          // Получаем значение EnableAutoCheck
          this.isTimerActive = setting.EnableAutoCheck === true;
          
          console.log(`Загружена настройка таймера: ${this.isTimerActive ? 'активен' : 'неактивен'}`);
        }
      } catch (error) {
        // Если списка настроек нет, используем значение по умолчанию
        console.log("Список настроек не найден, используем значение по умолчанию (таймер отключен)");
        this.isTimerActive = false;
      }
    } catch (error) {
      console.error("Ошибка при загрузке настроек таймера:", error);
      this.isTimerActive = false;
    }
  }

  /**
   * Сохраняет настройки таймера в список SharePoint
   */
  private async saveTimerSettings(): Promise<void> {
    try {
      // Проверяем существование списка настроек
      try {
        const settings = await this.sp.web.lists.getByTitle(this.SETTINGS_LIST_NAME).items
          .filter("Title eq 'TimerSettings'")();
        
        if (settings.length > 0) {
          // Обновляем существующую запись
          await this.sp.web.lists.getByTitle(this.SETTINGS_LIST_NAME).items
            .getById(settings[0].ID).update({
              EnableAutoCheck: this.isTimerActive
            });
          
          console.log(`Настройка таймера обновлена: ${this.isTimerActive ? 'активен' : 'неактивен'}`);
        } else {
          // Создаем новую запись
          await this.sp.web.lists.getByTitle(this.SETTINGS_LIST_NAME).items.add({
            Title: "TimerSettings",
            EnableAutoCheck: this.isTimerActive
          });
          
          console.log(`Настройка таймера создана: ${this.isTimerActive ? 'активен' : 'неактивен'}`);
        }
      } catch (error) {
        // Если список не существует, выводим сообщение
        console.error(`Ошибка при сохранении настроек: список ${this.SETTINGS_LIST_NAME} не существует.`);
        console.log(`Пожалуйста, создайте список ${this.SETTINGS_LIST_NAME} вручную и добавьте булево поле EnableAutoCheck.`);
      }
    } catch (error) {
      console.error("Ошибка при сохранении настроек таймера:", error);
    }
  }

  /**
   * Включает или выключает таймер проверки
   * @param isActive Флаг активности таймера
   */
  private toggleTimerActive(isActive: boolean): void {
    this.isTimerActive = isActive;
    
    // Сохраняем состояние в SharePoint списке
    this.saveTimerSettings().catch(error => {
      console.error('Ошибка при сохранении настроек таймера:', error);
    });
    
    // Если таймер должен быть остановлен
    if (!isActive && this.timerInterval) {
      clearInterval(this.timerInterval);
      this.timerInterval = null;
      console.log("Таймер остановлен пользователем");
      
      // Логируем остановку
      Logger.info(LOG_SOURCE, 'Автоматическая проверка деактивирована пользователем', {
        user: this.context.pageContext.user.displayName,
        timestamp: new Date().toISOString()
      }).catch(error => console.error('Ошибка при логировании:', error));
    } 
    // Если таймер должен быть запущен
    else if (isActive && !this.timerInterval) {
      this.startTimer();
      console.log("Таймер запущен пользователем");
      
      // Логируем запуск
      Logger.info(LOG_SOURCE, 'Автоматическая проверка активирована пользователем', {
        user: this.context.pageContext.user.displayName,
        timestamp: new Date().toISOString()
      }).catch(error => console.error('Ошибка при логировании:', error));
    }
  }

  /**
   * Запуск таймера для периодической проверки
   */
  private startTimer(): void {
    // Проверяем, должен ли таймер быть активен
    if (!this.isTimerActive) {
      console.log("Таймер отключен согласно настройкам");
      return;
    }
    
    // Проверяем, не запущен ли уже таймер
    if (this.timerInterval) {
      console.log("Таймер уже запущен");
      return;
    }
    
    // Запускаем таймер каждые 5 минут (300000 миллисекунд)
    this.timerInterval = setInterval(() => {
      this.checkListForChanges().catch(error => {
        console.error('Ошибка при проверке списка:', error);
      });
    }, 300000);
    
    console.log("Таймер запущен для проверки каждые 5 минут");
    
    // Выполним первичную проверку сразу
    this.checkListForChanges().catch(error => {
      console.error('Ошибка при первичной проверке списка:', error);
    });
  }

  /**
   * Проверка триггер-списка на наличие элементов для обработки
   */
  private async checkListForChanges(): Promise<void> {
    try {
      await Logger.info(LOG_SOURCE, "Проверка триггер-списка на предмет изменений", {
        triggerList: this.TRIGGER_LIST_NAME,
        timestamp: new Date().toISOString()
      });
      
      try {
        // Проверяем существование списка перед запросом элементов
        await this.sp.web.lists.getByTitle(this.TRIGGER_LIST_NAME).select('Title')();
        
        // Получаем все элементы из триггер-списка со статусом New
        const items = await this.sp.web.lists.getByTitle(this.TRIGGER_LIST_NAME).items
          .filter("Status eq 'New' and Trigger eq 1")();
        
        if (items.length === 0) {
          await Logger.info(LOG_SOURCE, "Элементы для обработки не найдены");
          return;
        }
        
        // Обрабатываем найденные элементы
        for (const item of items) {
          // Получаем параметры из элемента
          const requestedListName = item.ListName || this.properties.listName || "TestTasks";
          const actualListName = this.getActualListName(requestedListName);
          
          // Создаем объект параметров для обработки
          const requestBody: IApiRequestData = {
            listName: actualListName,
            urgent: item.Urgent === true,
            customer: item.Customer,
            region: item.Region,
            level: item.Level
          };
          
          try {
            // Запускаем бизнес-логику
            const result = await this.processListBusinessLogic(actualListName, requestBody);
            
            // Обновляем статус элемента
            await this.sp.web.lists.getByTitle(this.TRIGGER_LIST_NAME).items.getById(item.Id).update({
              Status: "Processed",
              ProcessedDate: new Date().toISOString(),
              ProcessedItems: result.length.toString()
            });
            
            // Обновляем дату последнего запуска
            const currentDate = new Date().toLocaleString();
            localStorage.setItem(this.LOCAL_STORAGE_KEY, currentDate);
            
            // Обновляем UI
            this._renderPlaceHolders();
            
            await Logger.info(LOG_SOURCE, "Элемент успешно обработан", {
              itemId: item.Id,
              processedItems: result.length
            });
          } catch (processError) {
            // Обрабатываем ошибку
            await Logger.error(LOG_SOURCE, `Ошибка при обработке элемента: ${processError}`, {
              itemId: item.Id,
              error: processError instanceof Error ? processError.message : String(processError)
            });
            
            // Обновляем статус элемента
            await this.sp.web.lists.getByTitle(this.TRIGGER_LIST_NAME).items.getById(item.Id).update({
              Status: "Error",
              ErrorMessage: processError instanceof Error ? processError.message : String(processError)
            });
          }
        }
      } catch (listError) {
        // Список не существует, выводим информацию в лог
        await Logger.warning(LOG_SOURCE, `Триггер-список не существует: ${listError}`, {
          triggerList: this.TRIGGER_LIST_NAME
        });
      }
    } catch (error) {
      // Общая ошибка
      await Logger.error(LOG_SOURCE, `Ошибка при проверке списка: ${error}`, {
        error: error instanceof Error ? error.message : String(error)
      });
      throw error; // Пробрасываем ошибку для обработки в вызывающем коде
    }
  }

  /**
   * Отрисовка элементов пользовательского интерфейса
   */
  private _renderPlaceHolders(): void {
    try {
      // Проверяем, отрисованы ли уже плейсхолдеры
      if (!this._topPlaceholder) {
        // Получаем плейсхолдер для верхней части страницы
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose.bind(this) }
        );

        // Если плейсхолдер существует
        if (this._topPlaceholder) {
          // Получаем дату последнего запуска из localStorage
          const lastRunDate = localStorage.getItem(this.LOCAL_STORAGE_KEY) || 'Еще не запускался';
          
          // Создаем элемент панели уведомлений
          const element: HTMLElement = document.createElement('div');
          element.className = styles.app;

          // Задаем содержимое панели с добавлением переключателя
          element.innerHTML = `
            <div class="${styles.topPanel}">
              <div class="${styles.topPanelContent}">
                <div class="${styles.topPanelText}">
                  KPF A-Plus Сервис | Последний запуск: ${lastRunDate}
                </div>
                <div class="${styles.topPanelControls}">
                  <label class="${styles.toggleLabel}">
                    <input type="checkbox" id="toggleTimerCheckbox" ${this.isTimerActive ? 'checked' : ''} />
                    Автопроверка
                  </label>
                  <button id="runApiButton" class="${styles.button}">Запустить обработку</button>
                </div>
              </div>
            </div>
          `;

          // Добавляем плейсхолдер в DOM
          this._topPlaceholder.domElement.appendChild(element);

          // Добавляем обработчик события для кнопки запуска
          const button = document.getElementById('runApiButton');
          if (button) {
            button.addEventListener('click', () => {
              this._runManualProcessing().catch(error => {
                console.error('Ошибка при запуске обработки:', error);
              });
            });
          }
          
          // Добавляем обработчик события для чекбокса
          const checkbox = document.getElementById('toggleTimerCheckbox');
          if (checkbox) {
            checkbox.addEventListener('change', (e) => {
              const isChecked = (e.target as HTMLInputElement).checked;
              this.toggleTimerActive(isChecked);
            });
          }
          
          // Логируем успешное отображение UI с обработкой Promise
          Logger.info(LOG_SOURCE, 'Пользовательский интерфейс отображен', {
            lastRunDate: lastRunDate,
            timerActive: this.isTimerActive
          }).catch(error => console.error('Ошибка при логировании:', error));
        }
      }
    } catch (error) {
      // Логируем ошибку при отрисовке UI с обработкой Promise
      Logger.error(LOG_SOURCE, `Ошибка при отрисовке UI: ${error}`, {
        stack: error instanceof Error ? error.stack : undefined
      }).catch(error => console.error('Ошибка при логировании:', error));
    }
  }

  /**
   * Запуск ручной обработки через кнопку в UI
   */
  private async _runManualProcessing(): Promise<void> {
    try {
      await Logger.info(LOG_SOURCE, 'Запуск ручной обработки', {
        user: this.context.pageContext.user.displayName
      });
      
      // Получаем имя списка из свойств или используем значение по умолчанию
      const requestedListName = this.properties.listName || "TestTasks";
      const actualListName = this.getActualListName(requestedListName);
      
      // Сообщаем пользователю о начале обработки
      await Dialog.alert(`Запуск обработки списка ${requestedListName}${requestedListName !== actualListName ? ` (будет использован список ${actualListName})` : ''}...`);
      
      // Создаем объект параметров для обработки
      const requestBody: IApiRequestData = {
        listName: actualListName
      };
      
      // Запускаем бизнес-логику
      const result = await this.processListBusinessLogic(actualListName, requestBody);
      
      // Обновляем дату последнего запуска
      const currentDate = new Date().toLocaleString();
      localStorage.setItem(this.LOCAL_STORAGE_KEY, currentDate);
      
      // Обновляем UI
      this._renderPlaceHolders();
      
      // Логируем успешное завершение
      await Logger.info(LOG_SOURCE, 'Обработка успешно завершена', { 
        requestedListName,
        actualListName,
        itemsProcessed: result.length,
        timestamp: new Date().toISOString()
      });
      
      // Показываем результат пользователю
      await Dialog.alert(`Обработка успешно завершена! 
      Обработано элементов: ${result.length}
      Запрошенный список: ${requestedListName}
      Фактический список: ${actualListName}`);
    } catch (error) {
      // Логируем ошибку
      await Logger.error(LOG_SOURCE, `Ошибка при ручной обработке: ${error}`, {
        operation: 'runManualProcessing',
        user: this.context.pageContext.user.displayName,
        properties: this.properties,
        timestamp: new Date().toISOString()
      });

      // Показываем ошибку пользователю
      await Dialog.alert(`Ошибка при обработке: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Получает реальное имя списка из карты маппинга
   * @param requestedListName Имя списка из запроса
   * @returns Реальное имя списка в SharePoint
   */
  private getActualListName(requestedListName: string): string {
    // Приводим к нижнему регистру для case-insensitive сравнения
    const lowerCaseRequestedName = requestedListName.toLowerCase();
    
    // Ищем маппинг в словаре (с учетом регистра)
    const mappedName = this.LIST_NAME_MAPPING[requestedListName];
    if (mappedName) {
      return mappedName;
    }
    
    // Ищем маппинг без учета регистра
    for (const key in this.LIST_NAME_MAPPING) {
      if (key.toLowerCase() === lowerCaseRequestedName) {
        return this.LIST_NAME_MAPPING[key];
      }
    }
    
    // Если маппинг не найден, возвращаем оригинальное имя
    return requestedListName;
  }

  /**
   * Метод с бизнес-логикой для обработки элементов списка
   * @param listName Реальное имя списка для обработки
   * @param requestData Полные данные запроса
   * @returns Массив обработанных элементов
   */
  private async processListBusinessLogic(listName: string, requestData: IApiRequestData): Promise<IProcessedItem[]> {
    try {
      await Logger.info(LOG_SOURCE, `Начало выполнения бизнес-логики для списка: ${listName}`, {
        requestData,
        timestamp: new Date().toISOString()
      });
      
      // Массив для хранения результатов обработки
      const results: IProcessedItem[] = [];
      
      try {
        // Пробуем получить доступ к списку для проверки его существования
        await this.sp.web.lists.getByTitle(listName).select('Title')();
        
        // Выбор стратегии обработки в зависимости от имени списка
        let items = [];
        switch (listName.toLowerCase()) {
          case 'testtasks':
            // Специальная логика для списка TestTasks
            items = await this.processTestTasksList(this.sp, listName);
            break;
            
          case 'orders':
            // Специальная логика для списка Orders
            items = await this.processOrdersList(this.sp, listName, requestData);
            break;
            
          case 'customers':
            // Специальная логика для списка Customers
            items = await this.processCustomersList(this.sp, listName, requestData);
            break;
            
          default:
            // Общая логика по умолчанию - просто обрабатываем элементы со статусом "New"
            items = await this.sp.web.lists.getByTitle(listName).items
              .filter("Status eq 'New'")();
            break;
        }
        
        await Logger.info(LOG_SOURCE, `Найдено элементов для обработки: ${items.length}`, {
          listName,
          timestamp: new Date().toISOString()
        });
        
        // Обработка каждого элемента
        for (const item of items) {
          try {
            await Logger.info(LOG_SOURCE, `Обработка элемента: ${item.ID}`, {
              listName,
              itemId: item.ID,
              itemTitle: item.Title || 'Без названия',
              timestamp: new Date().toISOString()
            });
            
            // Обновляем элемент
            await this.sp.web.lists.getByTitle(listName).items.getById(item.ID).update({
              ProcessedDate: new Date().toISOString(),
              Status: "Processed"
              // Другие поля для обновления
            });
            
            // Добавляем результат обработки
            results.push({
              id: item.ID,
              title: item.Title || 'Без названия',
              processed: true,
              timestamp: new Date().toISOString()
            });
            
            await Logger.info(LOG_SOURCE, `Элемент ${item.ID} успешно обработан`, {
              itemId: item.ID,
              timestamp: new Date().toISOString()
            });
          } catch (updateError) {
            // Логируем ошибку обновления элемента, но продолжаем с другими
            await Logger.error(LOG_SOURCE, `Ошибка при обновлении элемента ${item.ID}: ${updateError}`, {
              itemId: item.ID,
              error: updateError instanceof Error ? updateError.message : String(updateError)
            });
          }
        }
        
      } catch (listError) {
        // Обрабатываем ошибку доступа к списку
        await Logger.error(LOG_SOURCE, `Ошибка при доступе к списку ${listName}: ${listError}`, {
          listName,
          error: listError instanceof Error ? listError.message : String(listError)
        });
        
        // Выбрасываем исключение, чтобы оно было обработано выше
        throw new Error(`Список '${listName}' не существует или недоступен. Проверьте имя списка и права доступа.`);
      }
      
      // Возвращаем результаты
      return results;
    } catch (error) {
      await Logger.error(LOG_SOURCE, `Ошибка при выполнении бизнес-логики: ${error}`, {
        listName,
        requestData,
        stack: error instanceof Error ? error.stack : undefined,
        timestamp: new Date().toISOString()
      });
      throw error;
    }
  }

  /**
   * Обработка списка TestTasks
   * @param sp Инициализированный объект PnP SP
   * @param listName Имя списка
   * @returns Массив элементов для обработки
   */
  private async processTestTasksList(sp: SPFI, listName: string): Promise<Record<string, unknown>[]> {
    try {
      // Специальная логика для списка TestTasks
      // Сначала проверим, существуют ли элементы с фильтром "Status eq 'New'"
      const query = "Status eq 'New'";
      
      // Попробуем получить элементы
      const items = await sp.web.lists.getByTitle(listName).items
        .filter(query)();
        
      await Logger.info(LOG_SOURCE, `Обработка TestTasks: получено ${items.length} элементов`, {
        filter: query,
        listName
      });
      
      return items;
    } catch (error) {
      // Более подробное логирование ошибки
      await Logger.error(LOG_SOURCE, `Ошибка при получении элементов из списка ${listName}: ${error}`, {
        listName,
        errorDetails: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined
      });
      
      // Вернем пустой массив вместо ошибки
      return [];
    }
  }

  /**
   * Обработка списка Orders
   * @param sp Инициализированный объект PnP SP
   * @param listName Имя списка
   * @param requestData Данные запроса
   * @returns Массив элементов для обработки
   */
  private async processOrdersList(sp: SPFI, listName: string, requestData: IApiRequestData): Promise<Record<string, unknown>[]> {
    // В зависимости от данных запроса можем использовать разные фильтры
    let filter = "Status eq 'New'";
    
    if (requestData.urgent === true) {
      filter = "Status eq 'New' and Priority eq 'High'";
    } else if (requestData.customer) {
      filter = `Status eq 'New' and Customer eq '${requestData.customer}'`;
    }
    
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(filter)();
      
    await Logger.info(LOG_SOURCE, `Обработка Orders: получено ${items.length} элементов`, {
      filter,
      requestData
    });
    return items;
  }

  /**
   * Обработка списка Customers
   * @param sp Инициализированный объект PnP SP
   * @param listName Имя списка
   * @param requestData Данные запроса
   * @returns Массив элементов для обработки
   */
  private async processCustomersList(sp: SPFI, listName: string, requestData: IApiRequestData): Promise<Record<string, unknown>[]> {
    // Специальная логика для списка Customers
    let filter = "Status eq 'Active'";
    
    if (requestData.region) {
      filter += ` and Region eq '${requestData.region}'`;
    }
    
    if (requestData.level) {
      filter += ` and CustomerLevel eq '${requestData.level}'`;
    }
    
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(filter)();
      
    await Logger.info(LOG_SOURCE, `Обработка Customers: получено ${items.length} элементов`, {
      filter,
      requestData
    });
    return items;
  }

  /**
   * Освобождение ресурсов при уничтожении компонента
   */
  private _onDispose(): void {
    // Останавливаем таймер
    if (this.timerInterval) {
      clearInterval(this.timerInterval);
      this.timerInterval = null;
      console.log("Таймер остановлен");
    }
    
    Logger.info(LOG_SOURCE, 'Расширение отключено/выгружено')
      .catch(error => console.error('Ошибка при логировании:', error));
    console.log('Освобождение ресурсов плейсхолдера.');
  }
}