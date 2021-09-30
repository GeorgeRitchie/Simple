using System;
using System.Collections.Generic;

using _base.Controller;
using _base.Model;

namespace _base
{
	/// <summary>
	/// Main Class with basic API
	/// </summary>
	public class MainExcelController
	{
		private ExcelFileManager excelFileManager;
		private readonly SendMail sender;

		private DataOfSendingMails_OfSelectedSheet sendingMails = null;

		/// <summary>
		/// 
		/// </summary>
		public MainExcelController()
		{
			sender = new SendMail();
		}

		/// <summary>
		/// Opens given excel file, gets all sheets and returns their names.
		/// </summary>
		/// <param name="ExcelFileName">Excel file further operations.</param>
		/// <returns>List of found sheets' names</returns>
		public List<string> GetWorkbookSheetsNames(string ExcelFileName)
		{
			if (IsAnyExcelFileAlreadyOpened())
			{
				excelFileManager.Dispose();
				excelFileManager = null;
			}

			excelFileManager = new ExcelFileManager(OpenExcelFile.Open(ExcelFileName));

			return excelFileManager.GetWorkbookSheetsNames();
		}

		private bool IsAnyExcelFileAlreadyOpened()
		{
			return excelFileManager != null;
		}

		/// <summary>
		/// Sets choosen sheet to get receivers' names by <see cref="GetReceiversNames"/> and create required data by <see cref="MakeMailObjectsFromReceiversData"/>
		/// </summary>
		/// <exception cref="ArgumentException">This exception will be thrown if parameter SheetName is <see cref="null"/>, empty ("") or is not given by <see cref="GetWorkbookSheetsNames"/></exception>
		/// <param name="SheetName">Choosen sheet name. Sheet's name must be one of names given by <see cref="GetWorkbookSheetsNames"/>.</param>
		public void SetChoosenSheetName(string SheetName)
		{
			excelFileManager.SetChosenSheetName(SheetName);
		}

		/// <summary>
		/// Gets names of all receivers found in choosen sheet.
		/// </summary>
		/// <returns>Names of all receivers in choosen sheet</returns>
		public List<string> GetReceiversNames()
		{
			return excelFileManager.GetReceiversNames();
		}

		/// <summary>
		/// Sets selected receivers for creating required files and other data for mails in method <see cref="MakeMailObjectsFromReceiversData"/>.
		/// </summary>
		/// <param name="receivers">List of receivers that are choosen from total amount of receivers.</param>
		public void SetSelectedReceivers(List<string> receivers)
		{
			excelFileManager.SetSelectedReceivers(receivers);
		}

		/// <summary>
		/// Cleses excel application with all opened files.
		/// </summary>
		public void CloseApp()
		{
			excelFileManager.Dispose();
		}

		/// <summary>
		/// Creates files and other required data for mails of receivers.
		/// </summary>
		public void MakeMailObjectsFromReceiversData()
		{
			sendingMails = excelFileManager.MakeMailObjectsFromReceiversData();
		}

		/// <summary>
		/// Sends mails to receivers now. If sending mails are deffered sending, give as a parameter it's ID.
		/// If sending mails are simple sending, skip parameter.
		/// </summary>
		/// <exception cref="ArgumentException"></exception>
		/// <exception cref="InvalidOperationException"></exception>
		/// <exception cref="Exceptions.MailSendException"></exception>
		/// <param name="deferredMailID">ID of deferred mail sending</param>
		public void SendNow(long deferredMailID = -1)
		{
			if (IsDeferredSending(deferredMailID))
			{
				sender.SendDeferredSending(deferredMailID);
			}
			else
			{
				if (IsThereAnyMailToSend())
				{
					sender.SendNow(sendingMails);
				}
			}
		}

		private bool IsDeferredSending(long deferredMailID)
		{
			return deferredMailID > -1;
		}

		private bool IsThereAnyMailToSend()
		{
			return sendingMails != null;
		}

		/// <summary>
		/// Creates task for sending mails to receivers on specified Date and Time
		/// </summary>
		/// <exception cref="ArgumentException"></exception>
		/// <exception cref="InvalidOperationException"></exception>
		/// <param name="dateOfSending">Date and Time when mails should be send to receivers</param>
		public void SendLater(DateTime dateOfSending)
		{
			if (IsThereAnyMailToSend())
			{
				sendingMails.MailsSendDateTime = dateOfSending;

				sender.CreateDeferredSending(sendingMails);
			}
		}
	}
}
