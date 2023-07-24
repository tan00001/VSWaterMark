// <copyright file="OptionPageGrid.cs" company="Matt Lacey Limited">
// Copyright (c) Matt Lacey Limited. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace VSWaterMark
{
    public class OptionPageGrid : DialogPage
    {
        public const string CurrentFileName = "${currentfilename}";
        public const string CurrentDirectoryName = "${currentdirectoryname}";
        public const string CurrentProjectName = "${currentprojectname}";
        public const string CurrentFilePathInProject = "${currentfilepathinproject}";

#pragma warning disable SA1309 // Field names should not begin with underscore
        private readonly HashSet<string> _folders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private int _hashValue;
        private double _textSize = 16.0D;
        private string _fontFamilyName = "Consolas";
        private bool _isFontBold = false;
        private bool _isFontItalic = false;
        private bool _isFontUnderline = false;
        private bool _isFontStrikeThrough = false;
        private bool _usingImage = false;
        private string _displayedText = "To change this text, go to Tools > Options > Water Mark.";
        private string _displayedTextFormat = string.Empty;
        private bool _useCurrentFileName = false;
        private bool _useCurrentDirectoryName = false;
        private bool _useCurrentProjectName = false;
        private bool _useCurrentFilePathInProject = false;
        private bool _usingReplacements = false;
        private string _imagePath = string.Empty;
        private string _textColor = "Red";
        private string _borderColor = "Gray";
        private string _backgroundColor = "White";
        private double _borderMargin = 10D;
        private double _borderPadding = 3D;
        private double _borderOpacity = 0.7D;
#pragma warning restore SA1309 // Field names should not begin with underscore

        [Category("WaterMark")]
        [DisplayName("Enabled")]
        [Description("Show the watermark.")]
        public bool IsEnabled { get; set; } = true;

        [Category("WaterMark")]
        [DisplayName("Folders")]
        [Description("Show the watermark in these folders.")]
        public string Folders
        {
            get
            {
                return string.Join(";", _folders);
            }

            set
            {
                _folders.Clear();
                _folders.UnionWith(value.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));
            }
        }

        [Category("WaterMark")]
        [DisplayName("Top")]
        [Description("Show the watermark at the top.")]
        public bool PositionTop
        {
            get; set;
        }

        [Category("WaterMark")]
        [DisplayName("Left")]
        [Description("Show the watermark on the left.")]
        public bool PositionLeft
        {
            get; set;
        }

        [Category("Text")]
        [DisplayName("Displayed text")]
        [Description("The text to show in the watermark.")]
        public string DisplayedText
        {
            get
            {
                return _displayedText;
            }

            set
            {
                _displayedText = value ?? string.Empty;
                _usingImage = value.StartsWith("IMG:") == true;
                if (_usingImage)
                {
                    _imagePath = _displayedText.Substring(4);
                }
                else
                {
                    _imagePath = string.Empty;
                }

                var lowerCaseDisplayedText = _displayedText.ToLowerInvariant();
                (_useCurrentFileName, _displayedTextFormat) = Replace(_displayedText, lowerCaseDisplayedText, 0, CurrentFileName);
                (_useCurrentDirectoryName, _displayedTextFormat) = Replace(_displayedTextFormat, lowerCaseDisplayedText, 0, CurrentDirectoryName);
                (_useCurrentProjectName, _displayedTextFormat) = Replace(_displayedTextFormat, lowerCaseDisplayedText, 0, CurrentProjectName);
                (_useCurrentFilePathInProject, _displayedTextFormat) = Replace(_displayedTextFormat, lowerCaseDisplayedText, 0, CurrentFilePathInProject);
                _usingReplacements = _useCurrentFileName
                    || _useCurrentDirectoryName
                    || _useCurrentProjectName
                    || _useCurrentFilePathInProject;
            }
        }

        [Category("Text")]
        [DisplayName("Text size")]
        [Description("The size of the text in the watermark.")]
        public double TextSize
        {
            get
            {
                return _textSize;
            }

            set
            {
                _textSize = value;
                _hashValue = 0;
            }
        }

        [Category("Text")]
        [DisplayName("Font family")]
        [Description("The name of the font to use.")]
        public string FontFamilyName
        {
            get
            {
                return _fontFamilyName;
            }

            set
            {
                _fontFamilyName = value;
                _hashValue = 0;
            }
        }

        [Category("Text")]
        [DisplayName("Bold")]
        [Description("Should the text be displayed in bold.")]
        public bool IsFontBold
        {
            get
            {
                return _isFontBold;
            }

            set
            {
                _isFontBold = value;
                _hashValue = 0;
            }
        }

        [Category("Text")]
        [DisplayName("Italic")]
        [Description("Should the text be displayed in italic.")]
        public bool IsFontItalic
        {
            get
            {
                return _isFontItalic;
            }

            set
            {
                _isFontItalic = value;
                _hashValue = 0;
            }
        }

        [Category("Text")]
        [DisplayName("Underline")]
        [Description("Should the text be displayed with underline.")]
        public bool IsFontUnderline
        {
            get
            {
                return _isFontUnderline;
            }

            set
            {
                _isFontUnderline = value;
                _hashValue = 0;
            }
        }

        [Category("Text")]
        [DisplayName("Strike Through")]
        [Description("Should the text be displayed with strike through.")]
        public bool IsFontStrikeThrough
        {
            get
            {
                return _isFontStrikeThrough;
            }

            set
            {
                _isFontStrikeThrough = value;
                _hashValue = 0;
            }
        }

        [Category("Text")]
        [DisplayName("Color")]
        [Description("The color to use for the text. Can be a named value or Hex (e.g. '#FF00FF')")]
        public string TextColor
        {
            get
            {
                return _textColor;
            }

            set
            {
                _textColor = value;
                _hashValue = 0;
            }
        }

        [Category("Background")]
        [DisplayName("Border")]
        [Description("The color to use for the border. Can be a named value or Hex (e.g. '#FF00FF')")]
        public string BorderColor
        {
            get
            {
                return _borderColor;
            }

            set
            {
                _borderColor = value;
                _hashValue = 0;
            }
        }

        [Category("Background")]
        [DisplayName("Background")]
        [Description("The color to use for the background. Can be a named value or Hex (e.g. '#FF00FF')")]
        public string BackgroundColor
        {
            get
            {
                return _backgroundColor;
            }

            set
            {
                _backgroundColor = value;
                _hashValue = 0;
            }
        }

        [Category("Background")]
        [DisplayName("Margin")]
        [Description("Number of pixels between the border and the edge of the editor.")]
        public double BorderMargin
        {
            get
            {
                return _borderMargin;
            }

            set
            {
                _borderMargin = value;
            }
        }

        [Category("Background")]
        [DisplayName("Padding")]
        [Description("Number of pixels between the text and the border.")]
        public double BorderPadding
        {
            get
            {
                return _borderPadding;
            }

            set
            {
                _borderPadding = value;
                _hashValue = 0;
            }
        }

        [Category("Background")]
        [DisplayName("Opacity")]
        [Description("Strength of the background opacity.")]
        public double BorderOpacity
        {
            get
            {
                return _borderOpacity;
            }

            set
            {
                _borderOpacity = value;
                _hashValue = 0;
            }
        }

        public IReadOnlyCollection<string> GetContainingFolders()
        {
            return _folders;
        }

        public string GetImagePath() => _imagePath;

        public string GetDisplayTextFormat() => _displayedTextFormat;

        public int GetSettingsHash()
        {
            if (_hashValue != 0)
            {
                return _hashValue;
            }

            _hashValue = CombineHashCodes(
                _textSize.GetHashCode(),
                _fontFamilyName.GetHashCode(),
                _isFontBold.GetHashCode(),
                _isFontItalic.GetHashCode(),
                _isFontUnderline.GetHashCode(),
                _isFontStrikeThrough.GetHashCode(),
                _textColor?.GetHashCode() ?? 0,
                _borderColor?.GetHashCode() ?? 0,
                _backgroundColor?.GetHashCode() ?? 0,
                _borderMargin.GetHashCode(),
                _borderPadding.GetHashCode(),
                _borderOpacity.GetHashCode());

            return _hashValue;
        }

        public bool IsUsingImage() => _usingImage;

        public bool UsingReplacements() => _usingReplacements;

        public bool UseCurrentFileName() => _useCurrentFileName;

        public bool UseCurrentDirectoryName() => _useCurrentDirectoryName;

        public bool UseCurrentProjectName() => _useCurrentProjectName;

        public bool UseCurrentFilePathInProject() => _useCurrentFilePathInProject;

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            System.Diagnostics.Debug.WriteLine("OnClosed");

            Messenger.RequestUpdateAdornment();
        }

        private static int CombineHashCodes(params int[] hashCodes)
        {
            int combinedHash = 17;

            foreach (int hashCode in hashCodes)
            {
                combinedHash = (combinedHash * 31) + hashCode;
            }

            return combinedHash;
        }

        private static (bool, string) Replace(string text, string lowerCaseText, int startIndex, string lowerCaseReplacement)
        {
            var index = lowerCaseText.IndexOf(lowerCaseReplacement, startIndex);
            if (index < 0)
            {
                return (false, text.Substring(startIndex));
            }

            var tailIndex = index + lowerCaseReplacement.Length;

            (_, string textTail) = Replace(
                text,
                lowerCaseText,
                tailIndex,
                lowerCaseReplacement);

            text = text.Substring(startIndex, index) + lowerCaseReplacement + textTail;

            return (true, text);
        }
    }
}
