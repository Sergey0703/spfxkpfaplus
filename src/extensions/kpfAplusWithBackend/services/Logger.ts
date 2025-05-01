// src/extensions/kpfAplusWithBackend/services/Logger.ts

import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";

/**
 * Уровни логирования
 */
export enum LogLevel {
  Info = 1,
  Warning = 2,
  Error = 3,
  Critical = 4
}

/**
 * Интерфейс для деталей лога
 */
export interface ILogDetails {
  [key: string]: unknown;
}

/**
 * Сервис для логирования в SharePoint
 */
export class Logger {
  private static readonly LOG_LIST_NAME: string = "BackendLogs";
  private static context: ApplicationCustomizerContext;

  /**
   * Инициализирует логгер
   * @param context Контекст SPFx
   */
  public static initialize(context: ApplicationCustomizerContext): void {
    this.context = context;
  }

  /**
   * Логирует информационное сообщение
   * @param source Источник логирования
   * @param message Сообщение
   * @param details Дополнительные детали
   */
  public static async info(source: string, message: string, details?: ILogDetails): Promise<void> {
    await this.log(LogLevel.Info, source, message, details);
  }

  /**
   * Логирует предупреждение
   * @param source Источник логирования
   * @param message Сообщение
   * @param details Дополнительные детали
   */
  public static async warning(source: string, message: string, details?: ILogDetails): Promise<void> {
    await this.log(LogLevel.Warning, source, message, details);
  }

  /**
   * Логирует ошибку
   * @param source Источник логирования
   * @param error Объект ошибки или сообщение
   * @param details Дополнительные детали
   */
  public static async error(source: string, error: Error | string, details?: ILogDetails): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : error;
    const errorDetails: ILogDetails = error instanceof Error 
      ? { message: error.message, stack: error.stack, ...details }
      : details || {};
    
    await this.log(LogLevel.Error, source, errorMessage, errorDetails);
  }

  /**
   * Логирует критическую ошибку
   * @param source Источник логирования
   * @param error Объект ошибки или сообщение
   * @param details Дополнительные детали
   */
  public static async critical(source: string, error: Error | string, details?: ILogDetails): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : error;
    const errorDetails: ILogDetails = error instanceof Error 
      ? { message: error.message, stack: error.stack, ...details }
      : details || {};
    
    await this.log(LogLevel.Critical, source, errorMessage, errorDetails);
  }

  /**
   * Создает запись лога в списке SharePoint
   * @param level Уровень логирования
   * @param source Источник логирования
   * @param message Сообщение
   * @param details Дополнительные детали
   */
  private static async log(level: LogLevel, source: string, message: string, details?: ILogDetails): Promise<void> {
    // Логируем в консоль для режима разработки
    this.logToConsole(level, source, message, details);
    
    try {
      // Проверяем инициализацию контекста
      if (!this.context) {
        console.error("Logger не инициализирован. Вызовите Logger.initialize(context) перед использованием.");
        return;
      }

      // Инициализация PnP JS
      const sp = spfi().using(SPFx(this.context));
      
      // Проверяем существование списка
      const listExists = await this.checkListExists(sp);
      if (!listExists) {
        console.warn(`Список логов "${this.LOG_LIST_NAME}" не существует. Логи сохраняются только в консоли.`);
        return;
      }
      
      // Добавляем запись в список
      await sp.web.lists.getByTitle(this.LOG_LIST_NAME).items.add({
        Title: this.getLevelName(level),
        LogSource: source,
        LogMessage: message,
        LogDetails: details ? JSON.stringify(details) : "",
        LogDate: new Date(),
        LogUserName: this.context.pageContext.user.displayName,
        LogSeverity: level // Числовое значение от 1 до 4
      });
    } catch (error) {
      // Если не удалось записать в SharePoint, логируем в консоль
      console.error("Ошибка при сохранении лога в SharePoint:", error);
    }
  }

  /**
   * Проверяет существование списка логов
   * @param sp Объект PnP JS
   */
  private static async checkListExists(sp: SPFI): Promise<boolean> {
    try {
      await sp.web.lists.getByTitle(this.LOG_LIST_NAME).select("Title")();
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Логирует в консоль браузера
   * @param level Уровень логирования
   * @param source Источник логирования
   * @param message Сообщение
   * @param details Дополнительные детали
   */
  private static logToConsole(level: LogLevel, source: string, message: string, details?: ILogDetails): void {
    const timestamp = new Date().toISOString();
    const levelName = this.getLevelName(level);
    
    switch (level) {
      case LogLevel.Info:
        console.log(`[${timestamp}] [${levelName}] [${source}] ${message}`, details || "");
        break;
      case LogLevel.Warning:
        console.warn(`[${timestamp}] [${levelName}] [${source}] ${message}`, details || "");
        break;
      case LogLevel.Error:
      case LogLevel.Critical:
        console.error(`[${timestamp}] [${levelName}] [${source}] ${message}`, details || "");
        break;
    }
  }

  /**
   * Получает название уровня логирования
   * @param level Уровень логирования
   */
  private static getLevelName(level: LogLevel): string {
    switch (level) {
      case LogLevel.Info:
        return "Info";
      case LogLevel.Warning:
        return "Warning";
      case LogLevel.Error:
        return "Error";
      case LogLevel.Critical:
        return "Critical";
      default:
        return "Unknown";
    }
  }
}