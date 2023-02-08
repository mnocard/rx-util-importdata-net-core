using System;
using System.Collections.Generic;
using System.Globalization;
using NLog;
using ImportData.IntegrationServicesClient.Models;
using System.IO;
using System.Linq;

namespace ImportData
{
  public class Contract : Entity
  {
    public int PropertiesCount = 19;
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

      variableForParameters = this.Parameters[shift + 2].Trim();
      var counterparty = BusinessLogic.GetEntityWithFilter<ICounterparties>(c => c.Name == variableForParameters, exceptionList, logger);

      if (counterparty == null)
      {
        var message = string.Format("Не найден контрагент \"{0}\".", this.Parameters[shift + 2]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 3].Trim();
      var documentKind = BusinessLogic.GetEntityWithFilter<IDocumentKinds>(d => d.Name == variableForParameters, exceptionList, logger);

      if (documentKind == null)
      {
        var message = string.Format("Не найден вид документа \"{0}\".", this.Parameters[shift + 3]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 4].Trim();
      var contractCategory = BusinessLogic.GetEntityWithFilter<IContractCategories>(c => c.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 4].ToString()))
      {
        if (contractCategory == null)
        {
          var message = string.Format("Не найдена категория договора \"{0}\".", this.Parameters[shift + 4]);
          exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
          logger.Error(message);

          return exceptionList;
        }
      }

      var subject = this.Parameters[shift + 5];

      variableForParameters = this.Parameters[shift + 6].Trim();
      var businessUnit = BusinessLogic.GetEntityWithFilter<IBusinessUnits>(b => b.Name == variableForParameters, exceptionList, logger);

      if (businessUnit == null)
      {
        var message = string.Format("Не найдена НОР \"{0}\".", this.Parameters[shift + 6]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 7].Trim();
      IDepartments department = null;
      if (businessUnit != null)
        department = BusinessLogic.GetEntityWithFilter<IDepartments>(d => d.Name == variableForParameters && 
        (d.BusinessUnit == null || d.BusinessUnit.Id == businessUnit.Id), exceptionList, logger, true);
      else
        department = BusinessLogic.GetEntityWithFilter<IDepartments>(d => d.Name == variableForParameters , exceptionList, logger);

      if (department == null)
      {
        var message = string.Format("Не найдено подразделение \"{0}\".", this.Parameters[shift + 7]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var filePath = this.Parameters[shift + 8];
      var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

      DateTimeOffset validFrom = DateTimeOffset.MinValue;
      try
      {
        validFrom = ParseDate(this.Parameters[shift + 9], style, culture);
      }
      catch (Exception)
      {
        var message = string.Format("Не удалось обработать значение в поле \"Действует с\" \"{0}\".", this.Parameters[shift + 9]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      DateTimeOffset validTill = DateTimeOffset.MinValue;
      try
      {
        validTill = ParseDate(this.Parameters[shift + 10], style, culture);
      }
      catch (Exception)
      {
        var message = string.Format("Не удалось обработать значение в поле \"Действует по\" \"{0}\".", this.Parameters[shift + 10]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var totalAmount = 0.0;

      if (!string.IsNullOrWhiteSpace(this.Parameters[shift + 11]) && !double.TryParse(this.Parameters[shift + 11].Trim(), style, culture, out totalAmount))
      {
        var message = string.Format("Не удалось обработать значение в поле \"Сумма\" \"{0}\".", this.Parameters[shift + 11]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 12].Trim();
      var currency = BusinessLogic.GetEntityWithFilter<ICurrency>(c => c.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 12].Trim()) && currency == null)
      {
        var message = string.Format("Не найдено соответствующее наименование валюты \"{0}\".", this.Parameters[shift + 12]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var lifeCycleState = BusinessLogic.GetPropertyLifeCycleState(this.Parameters[shift + 13]);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 13].Trim()) && lifeCycleState == null)
      {
        var message = string.Format("Не найдено соответствующее значение состояния \"{0}\".", this.Parameters[shift + 13]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 14].Trim();
      var responsibleEmployee = BusinessLogic.GetEntityWithFilter<IEmployees>(e => e.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 14].Trim()) && responsibleEmployee == null)
      {
        var message = string.Format("Не найден Ответственный \"{3}\". Договор: \"{0} {1} {2}\". ", regNumber, regDate.ToString(), counterparty, this.Parameters[shift + 14].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      variableForParameters = this.Parameters[shift + 15].Trim();
      var ourSignatory = BusinessLogic.GetEntityWithFilter<IEmployees>(e => e.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 15].Trim()) && ourSignatory == null)
      {
        var message = string.Format("Не найден Подписывающий \"{3}\". Договор: \"{0} {1} {2}\". ", regNumber, regDate.ToString(), counterparty, this.Parameters[shift + 15].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      var note = this.Parameters[shift + 16];

      variableForParameters = this.Parameters[shift + 17].Trim();
      int idDocumentRegisters = int.Parse(variableForParameters);
      var documentRegisters = BusinessLogic.GetEntityWithFilter<IDocumentRegisters>(r => r.Id == idDocumentRegisters, exceptionList, logger);
      if (documentRegisters == null)
      {
        var message = string.Format("Не найден Журнал регистрации по ИД: \"{3}\". Договор: \"{0} {1} {2}\". ", regNumber, regDate.ToString(), counterparty, this.Parameters[shift + 17].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);

        return exceptionList;
      }

      var regState = this.Parameters[shift + 18].Trim();

      try
      {
        var regDateBeginningOfDay = BeginningOfDay(regDate.UtcDateTime);
        var contract = BusinessLogic.GetEntityWithFilter<IContracts>(x => x.RegistrationNumber != null && 
            x.RegistrationNumber == regNumber &&
            x.RegistrationDate == regDateBeginningOfDay &&
            x.Counterparty.Id == counterparty.Id &&
            x.DocumentRegister == documentRegisters, exceptionList, logger, true);
        if (contract == null)
          contract = new IContracts();

        // Обязательные поля.
        contract.Name = fileNameWithoutExtension;
        contract.Created = DateTimeOffset.UtcNow;
        contract.Counterparty = counterparty;
        contract.DocumentKind = documentKind;
        contract.DocumentGroup = contractCategory;
        contract.Subject = subject;
        contract.BusinessUnit = businessUnit;
        contract.Department = department;
        contract.ValidFrom = validFrom != DateTimeOffset.MinValue ? validFrom : Constants.defaultDateTime;
        contract.ValidTill = validTill != DateTimeOffset.MinValue ? validTill : Constants.defaultDateTime;
        contract.TotalAmount = totalAmount;
        contract.Currency = currency;
        contract.ResponsibleEmployee = responsibleEmployee;
        contract.OurSignatory = ourSignatory;
        contract.Note = note;
                
        contract.DocumentRegister = documentRegisters;
        contract.RegistrationDate = regDate != DateTimeOffset.MinValue ? regDate.UtcDateTime : Constants.defaultDateTime;
        contract.RegistrationNumber = regNumber;
        if (!string.IsNullOrEmpty(contract.RegistrationNumber) && contract.DocumentRegister != null)
          contract.RegistrationState = BusinessLogic.GetRegistrationsState(regState);

        var createdContract = BusinessLogic.CreateEntity<IContracts>(contract, exceptionList, logger);

        if (!string.IsNullOrWhiteSpace(filePath))
          exceptionList.AddRange(BusinessLogic.ImportBody(createdContract, filePath, logger));

        var documentRegisterId = 0;

        if (ExtraParameters.ContainsKey("doc_register_id"))
          if (int.TryParse(ExtraParameters["doc_register_id"], out documentRegisterId))
            exceptionList.AddRange(BusinessLogic.RegisterDocument(contract, documentRegisterId, regNumber, regDate, Constants.RolesGuides.RoleContractResponsible, logger));
          else
          {
            var message = string.Format("Не удалось обработать параметр \"doc_register_id\". Полученное значение: {0}.", ExtraParameters["doc_register_id"]);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Error(message);

            return exceptionList;
          }

        // Дополнительно обновляем свойство Состояние, так как после установки регистрационного номера Состояние сбрасывается в значение "В разработке"
        if (!string.IsNullOrEmpty(lifeCycleState))
          createdContract = createdContract.UpdateLifeCycleState(createdContract, lifeCycleState);
      }
      catch (Exception ex)
      {
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = ex.Message });
        logger.Error(ex, ex.Message);
        return exceptionList;
      }

      return exceptionList;
    }
  }
}
