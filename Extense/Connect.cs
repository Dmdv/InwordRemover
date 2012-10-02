using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Extense
{

	#region Read me for Add-in installation and setup information.

	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the ExtenseSetup project, 
	// right click the project in the Solution Explorer, then choose install.

	#endregion

	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[Guid("52147D55-9D90-4FBB-B40C-F3EBC34B5B90"), ProgId("Extense.Connect")]
	public class Connect : Object, IDTExtensibility2
	{
		private const string PopupMenuTag = "TextProcessing";
		private const string MenuItemTag = "MenuItemTag";
		private const string PopupMenuCaption = "Text Processing";

		private readonly object _missing = Missing.Value;
		private object _addInInstance;
		private Application _application;
		private ApplicationPresenter _presenter;

		/// <summary>
		/// Implements the OnConnection method of the IDTExtensibility2 interface.
		/// Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param name="application">Root object of the host application.</param>
		/// <param name="connectMode"> Describes how the Add-in is being loaded.</param>
		/// <param name="addInInst">Object representing this Add-in.</param>
		/// <param name="custom"></param>
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_application = (Application) application;
			_presenter = new ApplicationPresenter(_application);
			_addInInstance = addInInst;

			if (connectMode != ext_ConnectMode.ext_cm_Startup)
			{
				OnStartupComplete(ref custom);
			}
		}

		/// <summary>
		/// Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		/// Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>
		/// Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		/// Receives notification that the host application has completed loading.
		/// </summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnStartupComplete(ref Array custom)
		{
			_application.WindowSelectionChange += _presenter.OnWindowSelectionChange;

			try
			{
				var popup = _presenter.FindPopupMenu(PopupMenuTag);
				if (popup == null)
				{
					popup =
						(CommandBarPopup)
						_presenter.ActiveMenuBar.Controls.Add(MsoControlType.msoControlPopup, _missing, _missing, 2, true);
					popup.Caption = PopupMenuCaption;
					popup.Tag = PopupMenuTag;
				}
				popup.Visible = true;

				var button = _presenter.FindMenu(MenuItemTag);
				if (button == null)
				{
					button = (CommandBarButton) popup.Controls.Add(MsoControlType.msoControlButton, _missing, _missing, 1, true);
					button.Style = MsoButtonStyle.msoButtonIconAndCaption;
					button.BeginGroup = true;
					button.Caption = "Remove selected text";
					button.FaceId = 500;
					button.Tag = MenuItemTag;
				}
				button.Click += _presenter.OnProcessText;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		/// <summary>
		/// Implements the OnDisconnection method of the IDTExtensibility2 interface.
		/// Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param name="disconnectMode">Describes how the Add-in is being unloaded.</param>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
			if (disconnectMode != ext_DisconnectMode.ext_dm_HostShutdown)
			{
				OnBeginShutdown(ref custom);
			}
		}

		/// <summary>
		/// Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
		/// Receives notification that the host application is being unloaded.
		/// </summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnBeginShutdown(ref Array custom)
		{
			Reset();
		}

		private void Reset()
		{
			var popup = _presenter.FindPopupMenu(PopupMenuTag);
			while (popup != null)
			{
				popup.Delete(false);
				popup = _presenter.FindPopupMenu(PopupMenuTag);
			}
		}
	}
}