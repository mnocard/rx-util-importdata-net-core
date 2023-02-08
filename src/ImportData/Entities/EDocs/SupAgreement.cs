using System;
using System.Collections.Generic;
using System.Globalization;
using NLog;
using ImportData.IntegrationServicesClient.Models;
using System.IO;

namespace ImportData
{
  class SupAgreement : Entity
  {
    public int PropertiesCount = 20;
    /// <summary>
    /// Получить наименование число запрашиваемых параметров.
    /// </summary>
    /// <returns>Число запрашиваемых параметров.</returns>
    public override int GetPropertiesCount()
    {
      return PropertiesCount;
    }

    /// <summary>
    /// Сохранение сущности в RX.
    /// </summary>
    /// <param name="shift">Сдвиг по горизонтали в XLSX документе. Необходим для обработки документов, составленных из элементов разных сущностей.</param>
    /// <param name="logger">Логировщик.</param>
    /// <returns>Число запрашиваемых параметров.</returns>
    public override IEnumerable<Structures.ExceptionsStruct> SaveToRX(Logger logger, bool supplementEntity, string ignoreDuplicates, int shift = 0)
    {
      var exceptionList = new List<Structures.ExceptionsStruct>();
      var variableForParameters = this.Parameters[shift + 0].Trim();

      var regNumber = this.Parameters[shift + 0];
      DateTimeOffset regDate = DateTimeOffset.MinValue;
      var style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;
      var culture = CultureInfo.CreateSpecificCulture("en-GB");
      try
      {
        regDate = ParseDate(this.Parameters[shift + 1], style, culture);
      }
      catch (Exception)
      {
        var message = string.Format("Не удалось обработать дату регистрации \"{0}\".", this.Parameters[shift + 1]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var regNumberLeadingDocument = this.Parameters[shift + 2];

      var regDateLeadingDocument = DateTimeOffset.MinValue;
      try
      {
        regDateLeadingDocument = ParseDate(this.Parameters[shift + 3], style, culture);
      }
      catch (Exception)
      {
        var message = string.Format("Не удалось обработать дату регистрации ведущего документа \"{0}\".", this.Parameters[shift + 3]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 4].Trim();
      var counterparty = BusinessLogic.GetEntityWithFilter<ICounterparties>(c => c.Name == variableForParameters, exceptionList, logger);

      if (counterparty == null)
      {
        var message = string.Format("Не найден контрагент \"{0}\".", this.Parameters[shift + 4]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 5].Trim();
      var documentKind = BusinessLogic.GetEntityWithFilter<IDocumentKinds>(c => c.Name == variableForParameters, exceptionList, logger);

      if (documentKind == null)
      {
        var message = string.Format("Не найден вид документа \"{0}\".", this.Parameters[shift + 5]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var subject = this.Parameters[shift + 6];

      variableForParameters = this.Parameters[shift + 7].Trim();
      var businessUnit = BusinessLogic.GetEntityWithFilter<IBusinessUnits>(c => c.Name == variableForParameters, exceptionList, logger);

      if (businessUnit == null)
      {
        var message = string.Format("Не найдена НОР \"{0}\".", this.Parameters[shift + 7]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 8].Trim();
      IDepartments department = null;
      if (businessUnit != null)
        department = BusinessLogic.GetEntityWithFilter<IDepartments>(d => d.Name == variableForParameters &&
        (d.BusinessUnit == null || d.BusinessUnit.Id == businessUnit.Id), exceptionList, logger, true);
      else
        department = BusinessLogic.GetEntityWithFilter<IDepartments>(d => d.Name == variableForParameters, exceptionList, logger);

      if (department == null)
      {
        var message = string.Format("Не найдено подразделение \"{0}\".", this.Parameters[shift + 8]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var filePath = this.Parameters[shift + 9];
      var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

      DateTimeOffset validFrom = DateTimeOffset.MinValue;
      try
      {
        validFrom = ParseDate(this.Parameters[shift + 10], style, culture);
      }
      catch (Exception)
      {
        var message = string.Format("Не удалось обработать значение в поле \"Действует с\" \"{0}\".", this.Parameters[shift + 10]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      DateTimeOffset validTill = DateTimeOffset.MinValue;
      try
      {
        validTill = ParseDate(this.Parameters[shift + 11], style, culture);
      }
      catch (Exception)
      {
        var message = string.Format("Не удалось обработать значение в поле \"Действует по\" \"{0}\".", this.Parameters[shift + 11]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var totalAmount = 0.0;

      if (!string.IsNullOrWhiteSpace(this.Parameters[shift + 12]) && !double.TryParse(this.Parameters[shift + 12].Trim(), style, culture, out totalAmount))
      {
        var message = string.Format("Не удалось обработать значение в поле \"Сумма\" \"{0}\".", this.Parameters[shift + 12]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 13].Trim();
      var currency = BusinessLogic.GetEntityWithFilter<ICurrency>(c => c.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 13].Trim()) && currency == null)
      {
        var message = string.Format("Не найдено соответствующее наименование валюты \"{0}\".", this.Parameters[shift + 13]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var lifeCycleState = BusinessLogic.GetPropertyLifeCycleState(this.Parameters[shift + 14]);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 14].Trim()) && lifeCycleState == null)
      {
        var message = string.Format("Не найдено соответствующее значение состояния \"{0}\".", this.Parameters[shift + 14]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);
        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 15].Trim();
      var responsibleEmployee = BusinessLogic.GetEntityWithFilter<IEmployees>(c => c.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 15].Trim()) && responsibleEmployee == null)
      {
        var message = string.Format("Не найден Ответственный \"{3}\". Доп. соглашение: \"{0} {1} {2}\". ", regNumber, regDate.ToString(), counterparty, this.Parameters[shift + 15].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      variableForParameters = this.Parameters[shift + 16].Trim();
      var ourSignatory = BusinessLogic.GetEntityWithFilter<IEmployees>(c => c.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 16].Trim()) && ourSignatory == null)
      {
        var message = string.Format("Не найден Подписывающий \"{3}\". Доп. соглашение: \"{0} {1} {2}\". ", regNumber, regDate.ToString(), counterparty, this.Parameters[shift + 16].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      var note = this.Parameters[shift + 17];

      variableForParameters = this.Parameters[shift + 18].Trim();
      int idDocumentRegisters = int.Parse(variableForParameters);
      var documentRegisters = BusinessLogic.GetEntityWithFilter<IDocumentRegisters>(r => r.Id == idDocumentRegisters, exceptionList, logger);

      if (documentRegisters == null)
      {
        var message = string.Format("Не найден Журнал регистрации по ИД: \"{3}\". Договор: \"{0} {1} {2}\". ", regNumber, regDate.ToString(), counterparty, this.Parameters[shift + 18].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);

        return exceptionList;
      }

      var leadingDocument = BusinessLogic.GetEntityWithFilter<IContracts>(d => d.RegistrationNumber == regNumberLeadingDocument && d.RegistrationDate.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'") == regDateLeadingDocument.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'"), exceptionList, logger);

      if (leadingDocument == null)
      {
        var message = string.Format("Доп.соглашение не может быть импортировано. Не найден ведущий документ с реквизитами \"Дата документа\" {0}, \"Рег. №\" {1} и \"Контрагент\" {2}.", regDateLeadingDocument.ToString("d"), regNumberLeadingDocument, counterparty.Name);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var regState = this.Parameters[shift + 19].Trim();

      try
      {
        var regDateBeginningOfDay = BeginningOfDay(regDate.UtcDateTime);
        var supAgreement = BusinessLogic.GetEntityWithFilter<ISupAgreements>(x => x.RegistrationNumber == regNumber &&
            x.RegistrationDate == regDateBeginningOfDay &&
            x.Counterparty.Id == counterparty.Id &&
            x.DocumentRegister == documentRegisters, exceptionList, logger);
        if (supAgreement == null)
          supAgreement = new ISupAgreements();

        supAgreement.Name = fileNameWithoutExtension;
        supAgreement.Created = DateTimeOffset.UtcNow;
        supAgreement.LeadingDocument = leadingDocument;
        supAgreement.Counterparty = counterparty;
        supAgreement.DocumentKind = documentKind;
        supAgreement.Subject = subject;
        supAgreement.BusinessUnit = businessUnit;
        supAgreement.Department = department;
        supAgreement.ValidFrom = validFrom != DateTimeOffset.MinValue ? validFrom : Constants.defaultDateTime;
        supAgreement.ValidTill = validTill != DateTimeOffset.MinValue ? validTill : Constants.defaultDateTime;
        supAgreement.TotalAmount = totalAmount;
        supAgreement.Currency = currency;
        supAgreement.LifeCycleState = lifeCycleState;
        supAgreement.ResponsibleEmployee = responsibleEmployee;
        supAgreement.OurSignatory = ourSignatory;
        supAgreement.Note = note;

        supAgreement.DocumentRegister = documentRegisters;
        supAgreement.RegistrationNumber = regNumber;
        supAgreement.RegistrationDate = regDate != DateTimeOffset.MinValue ? regDate.UtcDateTime : Constants.defaultDateTime;
        if (!string.IsNullOrEmpty(supAgreement.RegistrationNumber) && supAgreement.DocumentRegister != null)
          supAgreement.RegistrationState = BusinessLogic.GetRegistrationsState(regState);

        var createdSupAgreement = BusinessLogic.CreateEntity<ISupAgreements>(supAgreement, exceptionList, logger);

        if (!string.IsNullOrWhiteSpace(filePath))
          exceptionList.AddRange(BusinessLogic.ImportBody(createdSupAgreement, filePath, logger));

        var documentRegisterId = 0;

        if (ExtraParameters.ContainsKey("doc_register_id"))
          if (int.TryParse(ExtraParameters["doc_register_id"], out documentRegisterId))
            exceptionList.AddRange(BusinessLogic.RegisterDocument(supAgreement, documentRegisterId, regNumber, regDate, Constants.RolesGuides.RoleContractResponsible, logger));
          else
          {
            var message = string.Format("Не удалось обработать параметр \"doc_register_id\". Полученное значение: {0}.", ExtraParameters["doc_register_id"]);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Error(message);

            return exceptionList;
          }
      }
      catch (Exception ex)
      {
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = ex.Message });

        return exceptionList;
      }

      return exceptionList;
    }
  }
}
