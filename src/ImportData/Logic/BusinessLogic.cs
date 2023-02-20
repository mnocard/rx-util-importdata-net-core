﻿using System;
using System.Linq;
using System.Collections.Generic;
using NLog;
using System.IO;
using ImportData.IntegrationServicesClient;
using ImportData.IntegrationServicesClient.Models;
using Simple.OData.Client;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using ImportData.IntegrationServicesClient.Exceptions;

namespace ImportData
{
  public class BusinessLogic
  {
    public static IEnumerable<string> ErrorList;

    /// <summary>
    /// Чтение атрибута EntityName.
    /// </summary>
    /// <param name="t">Тип класса.</param>
    /// <returns>Значение атрибута EntityName.</returns>
    private static string PrintInfo(Type t)
    {
      Attribute[] attrs = Attribute.GetCustomAttributes(t);

      foreach (Attribute attr in attrs)
      {
        if (attr is EntityName)
        {
          EntityName a = (EntityName)attr;

          return a.GetName();
        }
      }

      return string.Empty;
    }

    /// <summary>
    /// Получение экземпляра клиента OData.
    /// </summary>
    /// <returns>ODataClient.</returns>
    /// <remarks></remarks>
    public static ODataClient InstanceOData()
    {
      return Client.Instance();
    }

    #region Работа с сервисом интеграции.
    /// <summary>
    /// Получение сущности по фильтру.
    /// </summary>
    /// <typeparam name="T">Тип сущности.</typeparam>
    /// <param name="expression">Условие фильтрации.</param>
    /// <param name="exceptionList">Список ошибок.</param>
    /// <param name="logger">Логгер</param>
    /// <returns>Сущность.</returns>
    public static T GetEntityWithFilter<T>(Expression<Func<T, bool>> expression, List<Structures.ExceptionsStruct> exceptionList, Logger logger, bool isExpand = false) where T : class
    {
      Expression<Func<T, bool>> condition = expression;
      var filter = new ODataExpression(condition);

      logger.Info(string.Format("Получение сущности {0}", PrintInfo(typeof(T))));

      try
      {
        var entities = Client.GetEntitiesByFilter<T>(filter, isExpand);

        if (entities.Count() > 1)
        {
          var message = string.Format("Найдено несколько записей типа сущности \"{0}\" с именем \"{1}\". Проверьте, что выбрана верная запись.", PrintInfo(typeof(T)), entities.FirstOrDefault().ToString());
          exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
          logger.Warn(message);
        }

        return entities.FirstOrDefault();
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("(Not Found)"))
          throw new FoundMatchesException("Проверьте коррекность адреса службы интеграции Directum RX.");

        if (ex.Message.Contains("(Unauthorized)"))
          throw new FoundMatchesException("Проверьте коррекность указанной учетной записи.");

      }
      return null;
    }

    /// <summary>
    /// Получение сущностей.
    /// </summary>
    /// <typeparam name="T">Тип сущности.</typeparam>
    /// <returns>Список сущностей.</returns>
    public static IEnumerable<T> GetEntities<T>() where T : class
    {
      try
      {
        var entities = Client.GetEntities<T>();
        return entities;
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("(Not Found)"))
          throw new FoundMatchesException("Проверьте коррекность адреса службы интеграции Directum RX.");

        if (ex.Message.Contains("(Unauthorized)"))
          throw new FoundMatchesException("Проверьте коррекность указанной учетной записи.");
      }
      return null;
    }

    /// <summary>
    /// Создать сущность.
    /// </summary>
    /// <typeparam name="T">Тип сущности.</typeparam>
    /// <param name="entity">Экземпляр сущности.</param>
    /// <returns>Созданная сущность.</returns>
    public static T CreateEntity<T>(T entity, List<Structures.ExceptionsStruct> exceptionList, Logger logger) where T : class
    {
      logger.Info(string.Format("Создание сущности {0}", PrintInfo(typeof(T))));
      try
      {
        var entities = Client.CreateEntity<T>(entity, logger);

        return entities;
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("(Not Found)"))
          throw new FoundMatchesException("Проверьте коррекность адреса службы интеграции Directum RX.");

        if (ex.Message.Contains("(Unauthorized)"))
          throw new FoundMatchesException("Проверьте коррекность указанной учетной записи.");

        throw;
      }
    }

    /// <summary>
    /// Обновить сущность.
    /// </summary>
    /// <typeparam name="T">Тип сущности.</typeparam>
    /// <param name="entity">Экземпляор сущности.</param>
    /// <returns>Обновленная сущность.</returns>
    public static T UpdateEntity<T>(T entity, List<Structures.ExceptionsStruct> exceptionList, Logger logger) where T : class
    {
      try
      {
        var entities = Client.UpdateEntity<T>(entity);

        logger.Info(string.Format("Тип сущности {0} обновлен.", PrintInfo(typeof(T))));

        return entities;
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("(Not Found)"))
          throw new FoundMatchesException("Проверьте коррекность адреса службы интеграции Directum RX.");

        if (ex.Message.Contains("(Unauthorized)"))
          throw new FoundMatchesException("Проверьте коррекность указанной учетной записи.");
      }
      return null;
    }
    #endregion

    #region Работа с документами.
    /// <summary>
    /// Импорт тела документа.
    /// </summary>
    /// <param name="edoc">Экземпляр документа.</param>
    /// <param name="pathToBody">Путь к телу документа.</param>
    /// <param name="logger">Логировщик.</param>
    /// <returns>Список ошибок.</returns>
    public static IEnumerable<Structures.ExceptionsStruct> ImportBody(IElectronicDocuments edoc, string pathToBody, Logger logger)
    {
      var exceptionList = new List<Structures.ExceptionsStruct>();
      logger.Info("Импорт тела документа");

      try
      {
        // GetExtension возвращает расширение в формате ".<расширение>". Убираем точку.
        var extention = Path.GetExtension(pathToBody).Replace(".", "");
        var associatedApplication = BusinessLogic.GetEntityWithFilter<IAssociatedApplications>(a => a.Extension == extention, exceptionList, logger);

        if (associatedApplication != null)
        {
          var createdVersion = edoc.CreateVersion(edoc.Name, associatedApplication);
          var lastVersion = edoc.LastVersion();
          lastVersion.Body.Value = new byte[0];

          if (!File.Exists(pathToBody))
          {
            var message = string.Format("Не найден файл по заданому пути: \"{0}\"", pathToBody);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Warn(message);

            return exceptionList;
          }

          lastVersion.Body.Value = File.ReadAllBytes(pathToBody);
          lastVersion.AssociatedApplication = associatedApplication;

          bool isFillBody = edoc.FillBody(lastVersion);
        }
        else
        {
          var message = string.Format("Не обнаружено соответствующее приложение-обработчик для файлов с расширением \"{0}\"", extention);
          exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
          logger.Warn(message);

          return exceptionList;
        }
      }
      catch (Exception ex)
      {
        var message = string.Format("Не удается создать тело документа. Ошибка: \"{0}\"", ex.Message);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Warn(message);

        return exceptionList;
      }

      return exceptionList;
    }

    /// <summary>
    /// Регистрация документа.
    /// </summary>
    /// <param name="edoc">Экземпляр документа.</param>
    /// <param name="documentRegisterId">ИД журнала регистрации.</param>
    /// <param name="regNumber">Рег. №</param>
    /// <param name="regDate">Дата регистрации.</param>
    /// <param name="logger">Логировщик.</param>
    /// <returns>Список ошибок.</returns>
    public static IEnumerable<Structures.ExceptionsStruct> RegisterDocument(IOfficialDocuments edoc, int documentRegisterId, string regNumber, DateTimeOffset regDate, Guid defaultRegistrationRoleGuid, Logger logger)
    {
      var exceptionList = new List<Structures.ExceptionsStruct>();

      IRegistrationGroups regGroup = null;

      // TODO Кэшировать.
      var documentRegister = BusinessLogic.GetEntityWithFilter<IDocumentRegisters>(d => d.Id == documentRegisterId, exceptionList, logger);

      if (documentRegister != null && regDate != null && !string.IsNullOrEmpty(regNumber))
      {
        edoc.RegistrationDate = regDate;
        edoc.RegistrationNumber = regNumber;
        edoc.DocumentRegister = documentRegister;
        edoc.RegistrationState = BusinessLogic.GetRegistrationsState("Registerd");
        regGroup = documentRegister.RegistrationGroup;
      }
      else
      {
        var message = string.Format("Не удалось найти соответствующий реестр с ИД \"{0}\".", documentRegisterId);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      try
      {
        BusinessLogic.UpdateEntity<IOfficialDocuments>(edoc, exceptionList, logger);
      }
      catch (Exception ex)
      {
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = ex.Message });
      }

      return exceptionList;
    }
    #endregion

    #region Словари с перечислениями.
    /// <summary>
    /// Получение состояние регистрации.
    /// </summary>
    /// <param name="registrationState">Наименование состояния регистрации.</param>
    /// <returns>Состояние регистрации.</returns>
    public static string GetRegistrationsState(string key)
    {
      Dictionary<string, string> RegistrationState = new Dictionary<string, string>
        {
            {"Зарегистрирован", "Registered"},
            {"Зарезервирован", "Reserved"},
            {"Не зарегистрирован", "NotRegistered"},
            {"", null}
        };

        try
        {
            return RegistrationState[key];
        }
        catch (KeyNotFoundException ex)
        {
            throw new WellKnownKeyNotFoundException(key, ex.Message, ex.InnerException);
        }
    }

    /// <summary>
    /// Получение ЖЦ документа.
    /// </summary>
    /// <param name="lifeCycleStateName">Наименование ЖЦ документа.</param>
    /// <returns>ЖЦ.</returns>
    public static string GetPropertyLifeCycleState(string key)
    {
      Dictionary<string, string> LifeCycleStates = new Dictionary<string, string>
        {
            {"В разработке", "Draft"},
            {"Действующий", "Active"},
            {"Аннулирован", "Obsolete"},
            {"Расторгнут", "Terminated"},
            {"Исполнен", "Closed"},
            {"", null}
        };

        try
        {
            return LifeCycleStates[key];
        }
        catch (KeyNotFoundException ex)
        {
            throw new WellKnownKeyNotFoundException(key, ex.Message, ex.InnerException);
        }
    }

    /// <summary>
    /// Получение пола.
    /// </summary>
    /// <param name="sexPropertyName">Наименование пола.</param>
    /// <returns>Экземпляр записи "Пол".</returns>
    public static string GetPropertySex(string key)
    {
      Dictionary<string, string> sexProperty = new Dictionary<string, string>
        {
            {"Мужской", "Male"},
            {"Женский", "Female"},
            {"Male", "Male"},
            {"Female", "Female"},
            {"", null}
        };

        try
        {
            return sexProperty[key];
        }
        catch (KeyNotFoundException ex)
        {
            throw new WellKnownKeyNotFoundException(key, ex.Message, ex.InnerException);
        }
    }
    #endregion

    #region Проверка валидации.
    /// <summary>
    /// Проверка введенного ОГРН по количеству символов.
    /// </summary>
    /// <param name="psrn">ОГРН.</param>
    /// <param name="nonresident">Нерезидент.</param>
    /// <returns>Пустая строка, если длина ОГРН в порядке.
    /// Иначе текст ошибки.</returns>
    public static string CheckPsrnLength(string psrn, bool nonresident)
    {
      if (string.IsNullOrWhiteSpace(psrn))
        return string.Empty;

      if (nonresident)
        return string.Empty;

      psrn = psrn.Trim();

      return System.Text.RegularExpressions.Regex.IsMatch(psrn, @"(^\d{13}$)|(^\d{15}$)") ? string.Empty : Constants.Resources.IncorrecPsrnLength;
    }

    /// <summary>
    /// Проверка введенного КПП по количеству символов.
    /// </summary>
    /// <param name="trrc">КПП.</param>
    /// <param name="nonresident">Нерезидент.</param>
    /// <returns>Пустая строка, если длина КПП в порядке.
    /// Иначе текст ошибки.</returns>
    public static string CheckTrrcLength(string trrc, bool nonresident)
    {
      if (string.IsNullOrWhiteSpace(trrc))
        return string.Empty;

      if (nonresident)
        return string.Empty;

      trrc = trrc.Trim();
     
      return System.Text.RegularExpressions.Regex.IsMatch(trrc, @"(^\d{9}$)") ? string.Empty : Constants.Resources.IncorrecTrrcLength;

    }

    /// <summary>
    /// Проверка введенного кода подразделения по количеству символов.
    /// </summary>
    /// <param name="codeDepartment">Код подразделения.</param>
    /// <returns>Пустая строка, если длина кода подразделения в порядке.
    /// Иначе текст ошибки.</returns>
    public static string CheckCodeDepartmentLength(string codeDepartment)
    {
      if (string.IsNullOrWhiteSpace(codeDepartment))
        return string.Empty;

      codeDepartment = codeDepartment.Trim();

      return codeDepartment.Length <= 10 ? string.Empty : Constants.Resources.IncorrecCodeDepartmentLength;
    }

    /// <summary>
    /// Проверка ИНН на валидность.
    /// </summary>
    /// <param name="tin">Строка с ИНН.</param>
    /// <param name="forCompany">Признак того, что проверяется ИНН для компании.</param>
    /// <returns>Текст ошибки. Пустая строка для верного ИНН.</returns>
    public static string CheckTin(string tin, bool forCompany, bool nonresident)
    {
      if (string.IsNullOrWhiteSpace(tin))
        return string.Empty;

      tin = tin.Trim();

      if (nonresident)
        return string.Empty;
      

      // Проверить содержание ИНН. Должен состоять только из цифр. (Bug 87755)
      if (!Regex.IsMatch(tin, @"^\d*$"))
          return Constants.Resources.NotOnlyDigitsTin;

      // Проверить длину ИНН. Для компаний допустимы ИНН длиной 10 или 12 символов, для персон - только 12.
      if (forCompany && tin.Length != 10 && tin.Length != 12)
        return Constants.Resources.CompanyIncorrectTinLength;

      if (!forCompany && tin.Length != 12)
        return Constants.Resources.PeopleIncorrectTinLength;

      // Проверить контрольную сумму.
      if (!CheckTinSum(tin))
        return Constants.Resources.NotValidTin;


      // Проверить значения первых 2х цифр на нули.
      // 1 и 2 цифры - код субъекта РФ (99 для межрегиональной ФНС для физлиц и ИП или код иностранной организации).
      if (tin.StartsWith("00"))
        return Constants.Resources.NotValidTinRegionCode;

      return string.Empty;
    }

    /// <summary>
    /// Проверка контрольной суммы ИНН. Вызывается из CheckTinSum.
    /// </summary>
    /// <param name="tin">Строка ИНН. Передавать ИНН длиной 10-12 символов.</param>
    /// <param name="coefficients">Массив коэффициентов для умножения.</param>
    /// <returns>True, если контрольная сумма сошлась.</returns>
    private static bool CheckTinSum(string tin, int[] coefficients)
    {
      var sum = 0;
      for (var i = 0; i < coefficients.Count(); i++)
        sum += (int)char.GetNumericValue(tin[i]) * coefficients[i];
      sum = (sum % 11) % 10;
      return sum == (int)char.GetNumericValue(tin[coefficients.Count()]);
    }

    /// <summary>
    /// Проверка контрольной суммы ИНН.
    /// </summary>
    /// <param name="tin">ИНН.</param>
    /// <returns>True, если контрольная сумма сошлась.</returns>
    /// <remarks>Информация по ссылке: http://ru.wikipedia.org/wiki/Идентификационный_номер_налогоплательщика.</remarks>
    private static bool CheckTinSum(string tin)
    {
      var coefficient10 = new int[] { 2, 4, 10, 3, 5, 9, 4, 6, 8 };
      var coefficient11 = new int[] { 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 };
      var coefficient12 = new int[] { 3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 };
      return tin.Length == 10 ? CheckTinSum(tin, coefficient10) : (CheckTinSum(tin, coefficient11) && CheckTinSum(tin, coefficient12));
    }

        /// <summary>
        /// Проверка введенного ОКПО по количеству символов.
        /// </summary>
        /// <param name="psrn">ОКПО.</param>
        /// <returns>Пустая строка, если длина ОКПО в порядке.
        /// Иначе текст ошибки.</returns>
    public static string CheckNceoLength(string nceo, bool nonresident)
    {
      if (string.IsNullOrWhiteSpace(nceo))
        return string.Empty;

      if (nonresident)
        return string.Empty;

      nceo = nceo.Trim();

      return System.Text.RegularExpressions.Regex.IsMatch(nceo, @"(^\d{8}$)|(^\d{10}$)") ? string.Empty : Constants.Resources.IncorrecNceoLength;
    }
    #endregion
  }
}
