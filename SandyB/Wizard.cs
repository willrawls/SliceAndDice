using System;

namespace MetX.SliceAndDice
{
		public class Wizard
		{

			public Office.CommandBarButton mcbAddinButton;
			public Office.CommandBarButton mcbEditButton;
			public Office.CommandBarButton mcbShortcutButton;
			public Office.CommandBarButton mcbChangeToButton;
			public Office.CommandBarButton mcbAltChangeToButton;
			public Office.CommandBarButton mcbFavoritesButton;
			public Office.CommandBarButton mcbExternalsButton;
			public As WithEvents;
			public As WithEvents;
			public As WithEvents;
			public As WithEvents;
			public As WithEvents;
			public As WithEvents;
			public As WithEvents;
			public  m_oUI;
			public frmMain m_oWindow;
			public VBIDE.VBE m_oVBInst;
			public CAssocArray PropertyStack;
			public  asaCustom;
			public bool HostedByVB;
			public VbFileAttribute Attr;
			public CInsertionInfo InsertionInfo;
			public VBControl CurControl;
			public VBControl NewControl;
			public CAssocArray asaVar;
			public string sChoices;
			public string sTemplate;
			public string sProgID;
			public string sLine;
			public string sLastClassName;
			public string sLastChoice;
			public string sCodeToInsert;
			public string sCodeToInsert2;
			public string sPropertyName;
			public bool bInTemplate;
			public string sTemplateDatabasePath;
			public Window CurWindow;
			public bool bFound;
			public bool bShown;
			public int lFirstButton;
			public short Cancel;

			public string TemplateDatabaseName
			{
				get
				{
					try
{
					TemplateDatabaseName = m_oUI.TemplateDatabaseName;
					
        }
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
				}

			}

			public string TemplateDatabasePath
			{
				get
				{
					try
{
					TemplateDatabasePath = sBefore(m_oUI.TemplateDatabaseName, lTokenCount(m_oUI.TemplateDatabaseName, gsBS), gsBS) + gsBS;
					
        }
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
				}

			}

			public string Version
			{
				get
				{
					Version = App.Major + gsP + App.Minor + gsP + App.Revision;
				}

			}

			public CTemplate InternalCurrentTemplate
			{
				get
				{
					InternalCurrentTemplate = m_oUI.InternalCurrentTemplate;
				}

			}

			public CTemplate CurrentTemplate
			{
				get
				{
					CurrentTemplate = m_oUI.CurrentTemplate;
				}

			}

			public CSliceAndDice SliceAndDice
			{
				get
				{
					SliceAndDice = m_oUI.SliceAndDice;
				}

			}

			public Object SandyWindow
			{
				get
				{
					SandyWindow = m_oUI;
				}

			}

			public string CurrentTemplateNameAndCategory
			{
				get
				{
					CurrentTemplateNameAndCategory = m_oUI.txtName.Text;
				}

			}

			public CSadCommands SoftCommands
			{
				get
				{
					try
{
					SoftCommands = m_oUI.Complete;
					
        }
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
				}

			}

			public bool FavoriteCalledFromIDE
			{
				get
				{
					if ( m_oWindow Is null )
				{
 return; // ???;
					FavoriteCalledFromIDE = m_oWindow.FavoriteCalledFromIDE;
				}

				set
				{
					if ( m_oWindow Is null )
				{
 return; // ???;
					m_oWindow.FavoriteCalledFromIDE = value;
				}

			}


			public As EnumFiles			{

				switch Attr;
				Case "ALIAS": Attr = vbAlias;
				Case "ARCHIVE": Attr = vbArchive;
				Case "DIRECTORY": Attr = vbDirectory;
				Case "HIDDEN": Attr = vbHidden;
				Case "READONLY": Attr = vbReadOnly;
				Case "SYSTEM": Attr = vbSystem;
				Case "VOLUME": Attr = vbVolume;
				Case }
				else
				{: Attr = vbNormal;
			}
			public  Evaluate			{
				Evaluate = m_oUI.Evaluate(sExpression, asaVar);
			}
			public As FileExists			{
				FileExists = modGeneral.FileExists(sFilename);
			}
			public void HandleKeyPress			{
				if ( Shift = 3 )
				{;
				switch KeyCode;
				Case 69                                   ' "E"xternals window;
				ExternalsHandler_Click null, false, false;
				KeyCode = 0;
				Shift = 0;
				Case 70                                   ' "F"avorites;
				FavoritesHandler_Click null, false, false;
				KeyCode = 0;
				Shift = 0;
				Case 83                                   ' "S"lice and Dice window;
				MenuHandler_Click null, false, false;
				KeyCode = 0;
				Shift = 0;
				Case }
				else
				{;
				MsgBox "Combination key Shift-Ctrl-" + KeyCode + " pressed";
				KeyCode = 0;
				Shift = 0;
			}
			public As JumpTo			{
				m_oUI.JumpTo sTemplateName, bRecordInHistory, bSyncCategoryList;
			}
			public void NewTemplate			{
				m_oUI.NewTemplate bAutoCreate, sTitle, sDefaultShortName, bJumpToAfterCreate;
			}
			public As sChooseColor			{
				sChooseColor = m_oUI.sChooseColor(sInitialColor);
			}
			public As sChooseFile			{
				sChooseFile = m_oUI.sChooseFile(sPath, sFilename, sFilter);
			}
			public void ShowSplashScreen()
			{
				try
{;
				frmSplash.DetermineRegistration;
				frmSplash.Show;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As GetATemplate			{
				try
{;
				GetATemplate = null;

				if ( StrComp(sCategory, gsSpecialLineItemDelimiter + "CURRENT " + UCase$(gsCategory) + gsSpecialLineItemDelimiter, vbTextCompare) = 0 )
				{;
				GetATemplate = m_oUI.SliceAndDice.Categorys(sGetToken(m_oUI.InternalCurrentTemplate, 1, gsCategoryTemplateDelimiter)).Templates(sTemplate);
				}
				else
				{;
				GetATemplate = m_oUI.SliceAndDice.Categorys(sCategory).Templates(sTemplate);
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As LogError			{
				LogError = modGeneral.LogError(sModuleName, sProcName, lError, sErrorMsg, Erl);
			}
			public As sFileContents			{
				sFileContents = modGeneral.sFileContents(sPathAndFilename);
			}
			public As sGetGUID			{
				sGetGUID = modGeneral.sGetGUID(sProgID);
			}
			public  sChoose			{
				sChoose = modGeneral.sChoose(sChoices, sDelimiter, sDefault);
			}
			public  sExtractToken			{
				sExtractToken = modGeneral.sExtractToken(sOrigStr, nToken, strDelim);
			}
			public As BrowseForFolder			{
				BrowseForFolder = modGeneral.BrowseForFolder(m_oUI.hwnd, sPrompt);
			}
			public As bUserSure			{
				bUserSure = modGeneral.bUserSure(sPrompt);
			}
			public As NextNegativeUnique()
			{
				NextNegativeUnique = modGeneral.NextNegativeUnique;
			}
			public As lTokenCount			{
				lTokenCount = modGeneral.lTokenCount(sAllTokens, sDelim);
			}
			public As nZ			{
				nZ = modGeneral.nZ(vData, sDefault);
			}
			public As sDenormalize			{
				sDenormalize = modGeneral.sDenormalize(sLine);
			}
			public As sGetToken			{
				sGetToken = modGeneral.sGetToken(sAllTokens, lToken, sDelim);
			}
			public As sAfter			{
				sAfter = modGeneral.sAfter(sAllTokens, lToken, sDelim);
			}
			public As sBefore			{
				sBefore = modGeneral.sBefore(sAllTokens, lToken, sDelim);
			}
			public As sExcept			{
				sExcept = modGeneral.sExcept(sAllTokens, lToken, sDelim);
			}
			public As sInsertSpaces			{
				sInsertSpaces = modGeneral.sInsertSpaces(sToInsertInto);
			}
			public As sNormalize			{
				sNormalize = modGeneral.sNormalize(sLine);
			}
			public As sReplace			{
				sReplace = Replace(sAll, sFind, sReplaceWith);
			}
			public As zn			{
				zn = modGeneral.zn(sData);
			}
			public As lFindToken			{
				lFindToken = modGeneral.lFindToken(sAllTokens, sTokenToFind, sDelimiter);
			}
			public As AddPopup			{
				// TODO: Rewrite try/catch and/or goto. EH_Wizard_AddPopup;
				static bInHereAlready As Boolean;
				if ( bInHereAlready )
				{
 return; // ???;
				bInHereAlready = true;

				if ( ! HostedByVB )
				{
 return; // ???     ' Shell App Override;

				if ( m_oVBInst.CommandBars(sMenu) Is null )
				{
;
				MsgBox "Hmm... There doesn't seem to be anywhere I can place the following " + gsSliceAndDice + " button on the (nonexistant) '" + sMenu + "' menu:" + gs2EOLTab + "With Caption: " + sCaption + gs2EOLTab + "At position: " + nBefore;
				return; // ???;
			}
			public As AddButton			{
				// TODO: Rewrite try/catch and/or goto. EH_Wizard_AddButton;
				static bInHereAlready As Boolean;
				if ( bInHereAlready )
				{
 return; // ???;
				bInHereAlready = true;

				if ( ! HostedByVB )
				{
 return; // ???     ' Shell App Override;

				if ( m_oVBInst.CommandBars(sMenu) Is null )
				{
;
				MsgBox "Hmm... There doesn't seem to be anywhere I can place the following " + gsSliceAndDice + " button on the (nonexistant) '" + sMenu + "' menu:" + gs2EOLTab + "With Caption: " + sCaption + gs2EOLTab + "At position: " + nBefore;
				return; // ???;
			}
			public void DeleteCurrentTextSelection()
			{
				m_oUI.DeleteCurrentTextSelection;
			}
			public As DetermineFirstLineInSelection()
			{
				DetermineFirstLineInSelection = m_oUI.DetermineFirstLineInSelection;
			}
			public As DetermineLastLineInSelection()
			{
				DetermineLastLineInSelection = m_oUI.DetermineLastLineInSelection;
			}
			public void DoInsertion			{
				m_oUI.DoInsertion asaV, sTemplateToInsert, bSkipDeclarations;
			}
			public As FillTemplateWithUserInput			{
				FillTemplateWithUserInput = m_oUI.FillTemplateWithUserInput(asaX, sToParse, sCodeToInsert, sMsgBoxTitle);
			}
			public As FindLastProcLine			{
				FindLastProcLine = m_oUI.FindLastProcLine(sProcName, lProcType);
			}
			public void GetCategoryAndName			{
				m_oUI.GetCategoryAndName sCategoryAndName, sCategory, sShortName;
			}
			public As GetCurrentTextSelection()
			{
				GetCurrentTextSelection = m_oUI.GetCurrentTextSelection;
			}
			public void GetProcAtLine			{
				m_oUI.GetProcAtLine lCurrentLine, sProcName, lProcType;
			}
			public As InsertTemplate			{

				InsertionInfo = new CInsertionInfo();

				if ( SoftVars Is null )
				{
;
				InsertionInfo.SoftVars = new CAssocArray();
				}
				else
				{;
				InsertionInfo.SoftVars = SoftVars;
			}
			public As sChooseDatabase			{
				sChooseDatabase = m_oUI.sChooseDatabase(sPath, sFilename);
			}
			public As sPropertyType			{
				sPropertyType = m_oUI.sPropertyType(sFieldType);
			}
			public void AltChangeToHandler_Click			{
				ChangeToHandler_Click CommandBarControl, handled, CancelDefault;
			}
			public void BarHandler_Click			{
				MenuHandler_Click CommandBarControl, handled, CancelDefault;
			}
			public void ChangeToHandler_Click			{
				if ( ! HostedByVB )
				{
 return;




				asaVar = new CAssocArray();
				try
{;

				ChangeToHandler_Click_Start_Over:;

				gbCancelInsertion = false;
				foreach( var CurControl in m_oVBInst.SelectedVBComponent.Designer.SelectedVBControls );
				if ( gbCancelInsertion )
				{
 Exit For;

				if ( ! m_oUI.SetInternalCurrentTemplate("Change from" + gsCategoryTemplateDelimiter + CurControl.ClassName) )
				{;
				if ( ! bUserSure("Please set up a 'Change to' " + gsTemplate + " named" + gsEolTab + "'Change from" + gsCategoryTemplateDelimiter + CurControl.ClassName + gsA + vbNewLine + "before using this function on the '" + .ClassName + "' type control." + gs2EOL + "Select Yes to create the new " + gsTemplate + gsP + vbNewLine + "Select No to abort.") )
				{;
				return;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ExternalsHandler_Click			{
				try
{;
				m_oUI.ShowExternalsMenu;

				handled = true;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void FavoritesHandler_Click			{
				try
{;
				m_oUI.FavoriteCalledFromIDE = true;
				m_oUI.ShowFavMenu;

				handled = true;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void IDTExtensibility_OnAddInsUpdate			{
			}
			public void IDTExtensibility_OnConnection			{
				try
{;
				String sLoadList;

				frmSplash.lblDLLsLoaded(1).Text = "0";
				if ( GetSetting(App.ProductName, gsLast, "Show Splash", true) )
				{;
				bShown = true;
				frmSplash.Show;
				frmSplash.Refresh;
				}
				else
				{;
				bShown = false;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void IDTExtensibility_OnDisconnection			{
				try
{;

				//  Make sure the edited Template (if one) is saved;
				m_oUI.SaveTemplate;
				m_oWindow.Visible = false;

				//  Remove buttons from VB5 ToolBars;
				mcbAddinButton.Delete;
				mcbEditButton.Delete;
				mcbShortcutButton.Delete;
				mcbChangeToButton.Delete;
				mcbAltChangeToButton.Delete;
				mcbFavoritesButton.Delete;
				mcbExternalsButton.Delete;

				//  Insure all external object references are released correctly;
				mcbAddinButton = null;
				mcbEditButton = null;
				mcbShortcutButton = null;
				mcbChangeToButton = null;
				mcbAltChangeToButton = null;
				mcbFavoritesButton = null;

				MenuHandler = null;
				BarHandler = null;
				ShortcutHandler = null;
				ChangeToHandler = null;
				AltChangeToHandler = null;
				FavoritesHandler = null;
				ExternalsHandler = null;

				//  Save settings for next time;
				SaveSetting App.ProductName, "Settings", "Exit after insert", IIf(m_oUI.ExitAfterInsert, "true", "false");
				SaveSetting App.ProductName, "Settings", "Last " + gsTemplate, m_oUI.CurrentTemplateNameAndCategory;
				SaveFormPosition m_oWindow;

				//  Destroy object references;
				m_oWindow.HideAllWindows true;
				m_oWindow.ShutdownDLLs;
				m_oWindow = null;
				m_oUI.Parent = null;
				m_oUI.DBClassGen = null;
				m_oUI.Form_Unload Cancel;
				Unload m_oUI;
				m_oUI = null;

				//  Disassociate external objects referenced in this object;
				m_oVBInst = null;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void IDTExtensibility_OnStartupComplete			{
			}
			public void MenuHandler_Click			{
				try
{;
				m_oWindow.Visible = ! m_oWindow.Visible;
				if ( m_oWindow.Visible )
				{;
				m_oWindow.SetFocus;
				m_oWindow.ZOrder;
				}
				else
				{;
				m_oWindow.HideAllWindows;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ShortcutHandler_Click			{
				MenuHandler_Click CommandBarControl, handled, CancelDefault;
			}
			public As SaveToFile			{
				SaveToFile = modGeneral.SaveToFile(sFilename, sContents);
			}
			public void BrowseTo			{
				modGeneral.BrowseTo sURL;
			}
			public void ShowExternalsMenu()
			{
				try
{;
				if ( ! m_oUI Is null )
				{
;
				m_oUI.ShowExternalsMenu;
				};
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ShowFavoritesMenu()
			{
				try
{;
				if ( ! m_oUI Is null )
				{
;
				m_oUI.ShowFavMenu;
				};
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ShowMainWindow()
			{
				try
{;
				if ( m_oWindow Is null )
				{
 return;

				m_oWindow.WindowState = 0;

				m_oWindow.Visible = true;
				if ( m_oWindow.Visible )
				{;
				m_oWindow.SetFocus;
				m_oWindow.ZOrder;
				};
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void HideWindows()
			{
				try
{;
				if ( m_oWindow Is null )
				{
 return;

				m_oWindow.Visible = false;
				if ( ! m_oWindow.Visible )
				{;
				m_oWindow.HideAllWindows;
				};
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
		}
}
