using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Extense
{
	public class ApplicationPresenter
	{
		private readonly Application _application;
		private List<Range> _paras = new List<Range>();
		private readonly object _missing = Missing.Value;

		public ApplicationPresenter(Application application)
		{
			_application = application;
		}

		public CommandBar ActiveMenuBar
		{
			get { return _application.CommandBars.ActiveMenuBar; }
		}

		public CommandBarPopup FindPopupMenu(string tag)
		{
			return (CommandBarPopup)ActiveMenuBar.FindControl(MsoControlType.msoControlPopup, _missing, tag, true, true);
		}

		public CommandBarButton FindMenu(string tag)
		{
			return (CommandBarButton)ActiveMenuBar.FindControl(MsoControlType.msoControlButton, _missing, tag, _missing, true);
		}

		public void OnProcessText(CommandBarButton ctrl, ref bool canceldefault)
		{
			var span = new Stopwatch();
			span.Start();

			_application.ScreenUpdating = false;
			var state = _application.ActiveWindow.View.ShowHiddenText;
			_application.ActiveWindow.View.ShowHiddenText = true;

			var changes = RemoveHidden() + RemoveFragments("{}");

			_application.ActiveWindow.View.ShowHiddenText = state;
			_application.ScreenUpdating = true;
			ClearSelection();

			span.Stop();
			var secs = span.ElapsedMilliseconds / 1000d;
			const string Msg = "Deleted {0} fragments\r\nTime spent: {1} sec";
			MessageBox.Show(string.Format(Msg, changes, secs), "Fragments operations");
		}

		public void OnWindowSelectionChange(Selection sel)
		{
			var range = _application.Selection.Range;
			if (range.Start == range.End)
				ClearSelection();

			DefineReplacementScope(sel);
		}

		private Document ActiveDocument
		{
			get { return _application.ActiveDocument; }
		}

		private int InvisibleCharsCount
		{
			get { return _application.ActiveDocument.Characters.Count; }
		}

		private int VisibleCharsCount
		{
			get { return _application.ActiveDocument.Range().ComputeStatistics(WdStatistic.wdStatisticCharactersWithSpaces); }
		}

		private int RemoveFragments(string delims)
		{
			if (string.IsNullOrEmpty(delims) || delims.Length != 2)
				throw new ArgumentException("delims");

			var count = 0;
			foreach (var finder in new List<Find>(_paras.Select(para => para.Find)))
			{
				finder.ClearFormatting();
				finder.MatchWildcards = true;
				finder.Text = string.Format(@"\{0}*\{1}", delims[0], delims[1]);
				finder.Replacement.Text = string.Empty;
				finder.MatchWildcards = true;
				count += finder.ReplaceAll();
			}
			return count;
		}

		private int RemoveHidden()
		{
			var count = 0;
			foreach (var finder in new List<Find>(_paras.Select(para => para.Find)))
			{
				finder.ClearFormatting();
				finder.Font.Hidden = 1;
				finder.Replacement.Text = string.Empty;
				count += finder.ReplaceAll();
			}
			return count;
		}

		private void DefineReplacementScope(Selection sel)
		{
			// Во избежание массового Ctlr-A, легко доступного для юзера.
			var range = _application.ActiveDocument.Range();
			if (range.Start == sel.Range.Start || range.End == sel.Range.End)
			{
				_paras.Add(range);
				return;
			}

			// Остальные добавлять по параграфам.
			foreach (Paragraph para in sel.Paragraphs)
			{
				// Устраняет необходимость поиска '\r'
				if (sel.Range.Start == sel.Range.End || _paras.Exists(Match(para.Range)))
					continue;

				_paras.Add(para.Range);
				Trace.WriteLine(string.Format("Added para: [{0}:{1}]", para.Range.Start, para.Range.End));
			}
		}

		private static Predicate<Range> Match(Range para)
		{
			return x =>
				   x.Start == para.Start &&
				   x.End == para.End;
		}

		private void ClearSelection()
		{
			_paras.Clear();
			_paras = new List<Range>();
			_application.Selection.Collapse();
			Trace.WriteLine("Clear Selection");
		}

		private void ActionHit(int current)
		{
			ActiveDocument.UndoClear();
		}
	}
}