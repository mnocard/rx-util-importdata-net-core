﻿using ImportData.IntegrationServicesClient.Models;
using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ImportData
{
  class CompanyDirective : Entity
  {
    public int PropertiesCount = 14;

    public override int GetPropertiesCount()
    {
      return PropertiesCount;
    }

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
      var documentKind = BusinessLogic.GetEntityWithFilter<IDocumentKinds>(d => d.Name == variableForParameters, exceptionList, logger);

      if (documentKind == null)
      {
        var message = string.Format("Не найден вид документа \"{0}\".", this.Parameters[shift + 2]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var subject = this.Parameters[shift + 3];

      variableForParameters = this.Parameters[shift + 4].Trim();
      var businessUnit = BusinessLogic.GetEntityWithFilter<IBusinessUnits>(u => u.Name == variableForParameters, exceptionList, logger);


      if (businessUnit == null)
      {
        var message = string.Format("Не найдена НОР \"{0}\".", this.Parameters[shift + 4]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 5].Trim();
      IDepartments department = null;
      if (businessUnit != null)
        department = BusinessLogic.GetEntityWithFilter<IDepartments>(d => d.Name == variableForParameters &&
        (d.BusinessUnit == null || d.BusinessUnit.Id == businessUnit.Id), exceptionList, logger, true);
      else
        department = BusinessLogic.GetEntityWithFilter<IDepartments>(d => d.Name == variableForParameters, exceptionList, logger);

      if (department == null)
      {
        var message = string.Format("Не найдено подразделение \"{0}\".", this.Parameters[shift + 5]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var filePath = this.Parameters[shift + 6];
      var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

      variableForParameters = this.Parameters[shift + 7].Trim();
      var assignee = BusinessLogic.GetEntityWithFilter<IEmployees>(e => e.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 7].Trim()) && assignee == null)
      {
        var message = string.Format("Не найден Исполнитель \"{2}\". Приказ: \"{0} {1}\". ", regNumber, regDate.ToString(), this.Parameters[shift + 7].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      variableForParameters = this.Parameters[shift + 8].Trim();
      var preparedBy = BusinessLogic.GetEntityWithFilter<IEmployees>(e => e.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 8].Trim()) && preparedBy == null)
      {
        var message = string.Format("Не найден Подготавливающий \"{2}\". Приказ: \"{0} {1}\". ", regNumber, regDate.ToString(), this.Parameters[shift + 8].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      variableForParameters = this.Parameters[shift + 9].Trim();
      var ourSignatory = BusinessLogic.GetEntityWithFilter<IEmployees>(e => e.Name == variableForParameters, exceptionList, logger);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 9].Trim()) && ourSignatory == null)
      {
        var message = string.Format("Не найден Подписывающий \"{2}\". Приказ: \"{0} {1}\". ", regNumber, regDate.ToString(), this.Parameters[shift + 9].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);
      }

      var lifeCycleState = BusinessLogic.GetPropertyLifeCycleState(this.Parameters[shift + 10]);

      if (!string.IsNullOrEmpty(this.Parameters[shift + 10].Trim()) && lifeCycleState == null)
      {
        var message = string.Format("Не найдено соответствующее значение состояния \"{0}\".", this.Parameters[shift + 10]);
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
        logger.Error(message);

        return exceptionList;
      }

      var note = this.Parameters[shift + 11];

      variableForParameters = this.Parameters[shift + 12].Trim();
      int idDocumentRegisters = int.Parse(variableForParameters);
      var documentRegisters = BusinessLogic.GetEntityWithFilter<IDocumentRegisters>(r => r.Id == idDocumentRegisters, exceptionList, logger);

      if (documentRegisters == null)
      {
        var message = string.Format("Приложение не может быть импортировано. Не найден журнал регистрации по ИД \"{0}\" ", this.Parameters[shift + 12].Trim());
        exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
        logger.Warn(message);

        return exceptionList;
      }

      var regState = this.Parameters[shift + 13].Trim();

      try
      {
        var regDateBeginningOfDay = BeginningOfDay(regDate.UtcDateTime);
        var companyDirective = BusinessLogic.GetEntityWithFilter<ICompanyDirective>(x => x.RegistrationNumber == regNumber && 
            x.RegistrationDate == regDateBeginningOfDay &&
            x.DocumentRegister == documentRegisters, exceptionList, logger);
        if (companyDirective == null)
          companyDirective = new ICompanyDirective();

        companyDirective.Name = fileNameWithoutExtension;
        companyDirective.Created = DateTimeOffset.UtcNow;
        companyDirective.Name = fileNameWithoutExtension;
        companyDirective.DocumentKind = documentKind;
        companyDirective.Subject = subject;
        companyDirective.BusinessUnit = businessUnit;
        companyDirective.Department = department;
        companyDirective.Assignee = assignee;
        companyDirective.PreparedBy = preparedBy;
        companyDirective.OurSignatory = ourSignatory;
        companyDirective.LifeCycleState = lifeCycleState;
        companyDirective.Note = note;

        companyDirective.DocumentRegister = documentRegisters;
        companyDirective.RegistrationDate = regDate != DateTimeOffset.MinValue ? regDate.UtcDateTime : Constants.defaultDateTime;
        companyDirective.RegistrationNumber = regNumber;
        if (!string.IsNullOrEmpty(companyDirective.RegistrationNumber) && companyDirective.DocumentRegister != null)
          companyDirective.RegistrationState = BusinessLogic.GetRegistrationsState(regState);

        var createdcompanyDirective = BusinessLogic.CreateEntity<ICompanyDirective>(companyDirective, exceptionList, logger);

        if (!string.IsNullOrWhiteSpace(filePath))
          exceptionList.AddRange(BusinessLogic.ImportBody(createdcompanyDirective, filePath, logger));

        var documentRegisterId = 0;

        if (ExtraParameters.ContainsKey("doc_register_id"))
          if (int.TryParse(ExtraParameters["doc_register_id"], out documentRegisterId))
            exceptionList.AddRange(BusinessLogic.RegisterDocument(companyDirective, documentRegisterId, regNumber, regDate, Constants.RolesGuides.RoleIncomingDocumentsResponsible, logger));
          else
          {
            var message = string.Format("Не удалось обработать параметр \"doc_register_id\". Полученное значение: {0}.", ExtraParameters["doc_register_id"]);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Error(message);

            return exceptionList;
          }
        
        // Дополнительно обновляем свойство Состояние, так как после установки регистрационного номера Состояние сбрасывается в значение "В разработке"
        if (!string.IsNullOrEmpty(lifeCycleState))
          createdcompanyDirective = createdcompanyDirective.UpdateLifeCycleState(createdcompanyDirective, lifeCycleState);
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
