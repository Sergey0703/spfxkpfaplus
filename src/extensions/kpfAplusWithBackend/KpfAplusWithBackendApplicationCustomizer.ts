// src/extensions/kpfAplusWithBackend/KpfAplusWithBackendApplicationCustomizer.ts

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
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
 * Application Customizer с REST API и пользовательским интерфейсом 
 */
export default class KpfAplusWithBackendApplicationCustomizer
  extends BaseApplicationCustomizer<IKpfAplusWithBackendApplicationCustomizerProperties> {

  // Плейсхолдер для отображения UI
  private _topPlaceholder: PlaceholderContent | undefined;
  
  // Локальный ключ для хранения информации о последнем запуске
  private readonly LOCAL_STORAGE_KEY = 'KPF_DailyService_LastRunDate';

  /**
   * Карта соответствия имен списков
   * Используется для маппинга имен списков из запроса в реальные имена списков SharePoint
   */
  private readonly LIST_NAME_MAPPING: { [key: string]: string } = {
    'Timetable': 'TestTasks',
    'Schedule': 'TestTasks',
    'Расписание': 'TestTasks',
    'Tasks': 'TestTasks',      // Добавляем маппинг для "Tasks", чтобы он тоже указывал на "TestTasks"
    'TestTasks': 'TestTasks'   // Добавляем маппинг для самого себя, чтобы "TestTasks" всегда оставался "TestTasks"
  };

  @override
  public async onInit(): Promise<void> {
    try {
      // Инициализируем логгер
      Logger.initialize(this.context);
      
      // Логируем начало инициализации
      await Logger.info(LOG_SOURCE, `Инициализация расширения KPF A-Plus`, {
        version: '1.0.0',
        userAgent: window.navigator.userAgent,
        user: this.context.pageContext.user.displayName,
        url: window.location.href
      });
      
      // Стандартное логирование (для совместимости)
      Log.info(LOG_SOURCE, `Инициализация ${strings.Title}`);

      // Регистрация REST API конечной точки
      await this.context.httpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/SPFx/route/register?name=kpfaplus/api`,
        SPHttpClient.configurations.v1
      );
      await Logger.info(LOG_SOURCE, 'REST API точка зарегистрирована', { 
        apiPath: 'kpfaplus/api' 
      });

      // Установка обработчиков для плейсхолдеров
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

      // Первая отрисовка UI
      this._renderPlaceHolders();
      
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
   * Отрисовка элементов пользовательского интерфейса
   */
  private _renderPlaceHolders(): void {
    try {
      // Проверяем, отрисованы ли уже плейсхолдеры
      if (!this._topPlaceholder) {
        // Получаем плейсхолдер для верхней части страницы
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose }
        );

        // Если плейсхолдер существует
        if (this._topPlaceholder) {
          // Получаем дату последнего запуска из localStorage
          const lastRunDate = localStorage.getItem(this.LOCAL_STORAGE_KEY) || 'Еще не запускался';
          
          // Создаем элемент панели уведомлений
          const element: HTMLElement = document.createElement('div');
          element.className = styles.app;

          // Задаем содержимое панели
          element.innerHTML = `
            <div class="${styles.topPanel}">
              <div class="${styles.topPanelContent}">
                <div class="${styles.topPanelText}">
                  KPF A-Plus Сервис | Последний запуск: ${lastRunDate}
                </div>
                <div class="${styles.topPanelButton}">
                  <button id="runApiButton" class="${styles.button}">Запустить обработку</button>
                </div>
              </div>
            </div>
          `;

          // Добавляем плейсхолдер в DOM
          this._topPlaceholder.domElement.appendChild(element);

          // Добавляем обработчик события для кнопки
          const button = document.getElementById('runApiButton');
          if (button) {
            button.addEventListener('click', () => {
              this._runManualProcessing().catch(error => {
                console.error('Ошибка при запуске обработки:', error);
              });
            });
          }
          
          // Логируем успешное отображение UI
          Logger.info(LOG_SOURCE, 'Пользовательский интерфейс отображен', {
            lastRunDate: lastRunDate
          }).catch(error => console.error('Ошибка при логировании:', error));
          
        }
      }
    } catch (error) {
      // Логируем ошибку при отрисовке UI
      Logger.error(LOG_SOURCE, `Ошибка при отрисовке UI: ${error}`, {
        stack: error instanceof Error ? error.stack : undefined
      }).catch(error => console.error('Ошибка при логировании:', error));
    }
  }

  /**
   * Запуск ручной обработки
   */
  private async _runManualProcessing(): Promise<void> {
    try {
      // Создаем тело запроса (явно указываем TestTasks)
      const requestBody: IApiRequestData = {
        listName: "TestTasks" // Явно указываем имя списка, с которым хотим работать
      };

      // Получаем реальное имя списка с использованием маппинга
      const requestedListName = requestBody.listName;
      const actualListName = this.getActualListName(requestedListName || '');

      // Логируем начало обработки
      await Logger.info(LOG_SOURCE, 'Ручной запуск обработки', { 
        requestedListName,
        actualListName,
        user: this.context.pageContext.user.displayName,
        timestamp: new Date().toISOString()
      });

      // Сообщаем пользователю о начале обработки
      await Dialog.alert(`Запуск обработки списка ${requestedListName}${requestedListName !== actualListName ? ` (будет использован список ${actualListName})` : ''}...`);

      // Выполняем бизнес-логику
      const result = await this.processListBusinessLogic(actualListName, requestBody);

      // Обновляем дату последнего запуска
      const currentDate = new Date().toLocaleString();
      localStorage.setItem(this.LOCAL_STORAGE_KEY, currentDate);

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
      
      // Обновляем UI
      this._renderPlaceHolders();
    } catch (error) {
      // Логируем ошибку
      await Logger.error(LOG_SOURCE, `Ошибка при обработке: ${error}`, {
        operation: 'runManualProcessing',
        user: this.context.pageContext.user.displayName,
        properties: this.properties,
        timestamp: new Date().toISOString()
      });

      // Показываем ошибку
      await Dialog.alert(`Ошибка при обработке: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Обработчик для API запроса на обработку списка
   * Будет доступен по URL: _api/kpfaplus/api/processList
   */
  @override
  public async onHttpRequest(url: string, init: RequestInit, response: Response): Promise<Response> {
    // Разбираем URL для определения запрашиваемого метода API
    const urlParts = url.toLowerCase().split('/');
    const apiPath = urlParts[urlParts.length - 1];

    // Обрабатываем запрос processList
    if (apiPath === 'processlist') {
      await Logger.info(LOG_SOURCE, `Получен API запрос на обработку списка`, { 
        url,
        method: init.method,
        timestamp: new Date().toISOString()
      });
      
      try {
        // Получаем параметры запроса
        let requestBody: IApiRequestData = {};
        if (init.body) {
          try {
            const bodyText = init.body.toString();
            requestBody = JSON.parse(bodyText);
            await Logger.info(LOG_SOURCE, `Параметры запроса API`, {
              ...requestBody,
              timestamp: new Date().toISOString()
            });
          } catch (e) {
            await Logger.error(LOG_SOURCE, `Ошибка парсинга JSON: ${e}`, { 
              bodyText: init.body.toString() 
            });
          }
        }
        
        // Получаем имя списка из запроса или используем TestTasks по умолчанию
        const requestedListName = requestBody.listName || 'TestTasks';
        
        // Получаем реальное имя списка с использованием маппинга
        const actualListName = this.getActualListName(requestedListName.toString());
        
        // Логируем информацию о реальном имени списка, если оно отличается
        if (requestedListName !== actualListName) {
          await Logger.info(LOG_SOURCE, `Переназначение списка`, {
            requestedListName,
            actualListName,
            timestamp: new Date().toISOString()
          });
        }
        
        // Выполняем бизнес-логику с реальным именем списка
        const result = await this.processListBusinessLogic(actualListName, requestBody);
        
        // Обновляем дату последнего запуска
        localStorage.setItem(this.LOCAL_STORAGE_KEY, new Date().toLocaleString());
        
        // Логируем успешное выполнение API запроса
        await Logger.info(LOG_SOURCE, `API запрос успешно выполнен`, { 
          requestedListName,
          actualListName,
          processedItems: result.length,
          timestamp: new Date().toISOString()
        });
        
        // Возвращаем результат
        return Promise.resolve(
          new Response(JSON.stringify({
            success: true,
            requestedList: requestedListName,
            actualList: actualListName,
            processedItems: result.length,
            items: result
          }), {
            headers: {
              'Content-Type': 'application/json'
            },
            status: 200
          })
        );
      } catch (error) {
        // Логируем ошибку API запроса
        await Logger.error(LOG_SOURCE, `Ошибка при обработке API запроса: ${error}`, {
          url,
          method: 'onHttpRequest',
          timestamp: new Date().toISOString(),
          stack: error instanceof Error ? error.stack : undefined
        });
        
        // Возвращаем ошибку
        return Promise.resolve(
          new Response(JSON.stringify({
            success: false,
            error: error instanceof Error ? error.message : String(error)
          }), {
            headers: {
              'Content-Type': 'application/json'
            },
            status: 500
          })
        );
      }
    }
    
    // Если запрос не распознан, передаем управление дальше
    return response;
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
      
      // Инициализация PnP JS для текущего контекста
      const sp = spfi().using(SPFx(this.context));
      
      // Массив для хранения результатов обработки
      const results: IProcessedItem[] = [];
      
      try {
        // Пробуем получить доступ к списку для проверки его существования
        await sp.web.lists.getByTitle(listName).select('Title')();
        
        // Выбор стратегии обработки в зависимости от имени списка
        let items = [];
        switch (listName.toLowerCase()) {
          case 'testtasks':
            // Специальная логика для списка TestTasks
            items = await this.processTestTasksList(sp, listName);
            break;
            
          case 'orders':
            // Специальная логика для списка Orders
            items = await this.processOrdersList(sp, listName, requestData);
            break;
            
          case 'customers':
            // Специальная логика для списка Customers
            items = await this.processCustomersList(sp, listName, requestData);
            break;
            
          default:
            // Общая логика по умолчанию - просто обрабатываем элементы со статусом "New"
            items = await sp.web.lists.getByTitle(listName).items
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
            await sp.web.lists.getByTitle(listName).items.getById(item.ID).update({
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
    Logger.info(LOG_SOURCE, 'Расширение отключено/выгружено')
      .catch(error => console.error('Ошибка при логировании:', error));
    console.log('Освобождение ресурсов плейсхолдера.');
  }
}