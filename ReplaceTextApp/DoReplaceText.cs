using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using tvToolbox;

namespace ReplaceText
{
    /// <summary>
    /// ReplaceText.exe
    ///
    /// Run this program. It will prompt to create a folder with profile file:
    ///
    ///     ReplaceText.exe.txt
    ///
    /// The profile will contain help (see -Help) as well as default options.
    ///
    /// Note: This software creates its own support files (including DLLs or
    ///       other EXEs) in the folder that contains it. It will prompt you
    ///       to create its own desktop folder when you run it the first time.
    /// </summary>


    public partial class DoReplaceText : Form
    {
        private bool            mbSpecialCharactersProfileInitDone;
        private bool            mbUseSpecialCharacters;
        private const string    mcsFileChangedKey                   = "-Changed";
        private const string    mcsFileUnchangedKey                 = "-Unchanged";
        private const string    mcsFoundTextKey                     = "-FoundIn";
        private string          msCurrentProfileAbsPathFile         = null;

        [DllImport( "kernel32.dll" )]
        static extern bool AttachConsole( int dwProcessId );
        private const int ATTACH_PARENT_PROCESS = -1;

        public DoReplaceText()
        {
            InitializeComponent();
        }

        protected tvProfile moSpecialCharactersProfile = new tvProfile()
        {
            // Special character keys must be lowercase.
            // Uppercase equivalents will be added below.

             {"\\b",  (char)7}
            ,{"\\e",  (char)27}
            ,{"\\f",  '\f'}
            ,{"\\n",  '\n'}
            ,{"\\q",  '"'}
            ,{"\\r",  '\r'}
            ,{"\\s",  '\''}
            ,{"\\t",  '\t'}
            ,{"\\v",  '\v'}
        };

        private void DoReplaceText_Load(object sender, EventArgs e)
        {
            this.Hide();

            string      lsErrorMessages         = null;
            tvProfile   loProfile               = null;
            bool        lbCopyResultsToSTDOUT   = false;
            bool        lbDisplayResults        = false;
            bool        lbNoPrompts             = false;
            string      lsDisplayResultsCommand = null;

            try
            {
                loProfile = new tvProfile(Environment.GetCommandLineArgs());
                if ( loProfile.bExit )
                {
                    this.Hide();
                    System.Windows.Forms.Application.Exit();
                    return;
                }

                bool            lbUseRegularExpressions = loProfile.bValue("-UseRegularExpressions", false);
                                mbUseSpecialCharacters  = loProfile.bValue("-UseSpecialCharacters", false);
                RegexOptions    leRegexOptions = RegexOptions.None;
                string          lsSubReplacementToken = loProfile.sValue("-SubToken", "{SubToken}");
                                // This must appear before the array load so that at least one item is always available.
                                loProfile.GetAdd("-OldSubValue", "One of many old 'sub replacement' values goes here.");
                string[]        lsOldSubReplacementArray = loProfile.sOneKeyArray("-OldSubValue");
                                loProfile.GetAdd("-NewSubValue", "One of many new 'sub replacement' values goes here.");
                                this.ReplaceSpecialCharacters(lsOldSubReplacementArray);
                string[]        lsNewSubReplacementArray = loProfile.sOneKeyArray("-NewSubValue");
                                this.ReplaceSpecialCharacters(lsNewSubReplacementArray);

                                if ( lsNewSubReplacementArray.Length != lsOldSubReplacementArray.Length )
                                    lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine)
                                            + "The number of -NewSubValue items must match the number of -OldSubValue items.";

                msCurrentProfileAbsPathFile = Path.GetFullPath(loProfile.sLoadedPathFile);

                loProfile.GetAdd("-Help",
                        @"
Introduction


This utility will take any files (presumably text) referenced as -Files and
replace old text referenced as -OldText with new text referenced as -NewText.

Notes:

New text substrings MUST correspond to old text substrings, one-to-one.

If the values of -NewText and -OldText are IDENTICAL and -OldText is found in a
file, the file will be identified with mcsFoundTextKey and it will remain unchanged.

    BE CAREFUL!!! If you intend to do searches using this technique,
                  -CaseInsensitiveSearch should be set False.
                  You should use the -SearchOnly switch instead.

If -SearchOnly=False and the values of -NewText and -OldText are not identical
and -OldText is found in a file, the file will be identified with mcsFileChangedKey.

If -SearchOnly=False and the values of -NewText and -OldText are not identical and
-OldText is NOT found in a file, the file will be identified with mcsFileUnchangedKey.


The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the ""Software""), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.


Command-Line Usage


Open this utility's profile file to see additional options available. It
is usually located in the same folder as ""{EXE}"" and has the same
name with "".txt"" added (see ""{INI}"").

Profile file options can be overridden with command-line arguments. The
keys for any ""-key=value"" pairs passed on the command-line must match
those that appear in the profile (with the exception of the ""-ini"" key).

For example, the following invokes the use of an alternative profile file:

    {EXE} -ini=NewProfile.txt

This tells the software to display all files changed and unchanged:

    {EXE} -DisplayResults


Author:  George Schiro (GeoCode@Schiro.name)

Date:    1/27/2005




Options and Features


The main options for this utility are listed below with their default values.
A brief description of each feature follows.

-CaseInsensitiveSearch=False

    Set this switch True and the original version plus both uppercase and
    lowercase versions of each given -OldText value (not mixed-case versions)
    will be replaced (see -OldText below).

-CopyResultsToSTDOUT=False

    Set this switch True and replacement results will be copied to standard
    output. This can be useful if the software is run by another process.

-DisplayResults=True

    Set this switch False and replacement results will not be displayed using
    the -DisplayResultsModule (see below).

-DisplayResultsModule=Notepad.exe

    This is the software used to display replacement results (see
    -DisplayResults) above.

-FetchSource=False

    Set this switch True to fetch the source code for this utility
    from the EXE. Look in the containing folder for a ZIP file with
    the full project sources.

-Files=""One of many 'files to replace' pathfile(s) specifications goes here.""

    This is the specification of files to be processed. It can include
    wildcards.

    Note: This key may appear any number of times in the profile.

-Help= SEE PROFILE FOR DEFAULT VALUE

    This help text.

-IgnoreNoFilesFound=False

    Set this switch True and no error pop-up will appear if no files are
    actually found to process.

-NewSubValue=""One of many new 'sub replacement' values goes here.""

    This is a new ""sub replacement"" value to replace the corresponding
    -OldSubValue (see below) within each -OldText value (see below).

    See -SubToken below for more details.

    Note: This key may appear any number of times in the profile.

-NewText=""One of many new text replacement substrings goes here.""

    This is a new text substring to replace a corresponding old text
    substring (see -OldText below) in all of the files given by the
    various -Files specifications (see above).

    Note: This key may appear any number of times in the profile.

-NoPrompts=False

    Set this switch True and all pop-up prompts will be suppressed. You must
    use this switch whenever the software is run via a server computer batch
    job or job scheduler (ie. where no user interaction is permitted).

-OldSubValue=""One of many old 'sub replacement' values goes here.""

    This is an old ""sub replacement"" value to replace the corresponding
    -NewSubValue (see above) within each -OldText value (see below).

    See -SubToken below for more details.

    Note: This key may appear any number of times in the profile.

-OldText=""One of many old text substrings to replace goes here.""

    This is an old text substring to be replaced by a corresponding
    new text substring (see -NewText above) in all of the files given
    by the various -Files specifications (see above).

    Note: This key may appear any number of times in the profile.

-RecurseSubdirectories=True

    Set this switch False and only the base directory found in each
    -Files specification (see above) will be searched for -OldText
    (see above). Otherwise, every file matching the -Files specifications
    found in every subdirectory from each base subdirectory onward will
    be searched.

-SaveProfile=True

    Set this switch False to prevent saving to the profile file by this
    software. This is not recommended since status information is written
    to the profile after each run.

-SaveSansCmdLine=True

    Set this switch False to leave the profile file untouched after a command-
    line has been passed to the EXE and merged with the profile. When true,
    everything but command-line keys will be saved. When false, not even status
    information will be written to the profile file (ie. ""{INI}"").

-SearchOnly=True

    Set this switch False and all files matching the specifications in
    the -Files parameters (see above) will be updated if they contain at
    least one -OldText value (see above). Otherwise, each matching file
    that contains at least one -OldText value will be displayed with the
    mcsFoundTextKey key.

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the
    profile file at startup in command-line format. This may be helpful as a
    diagnostic.

-SubToken={SubToken}

    A ""sub replacement token"" (lsSubReplacementToken) can be used to
    pass a common substring value referenced as (-OldSubValue) to be
    replaced with the corresponding -NewSubValue (see above) within each
    (-OldText) value.

    Any number of -OldSubValue,-NewSubValue pairs can be given, all of
    which will be replaced in every -OldText value (if found there).

    This feature is useful if you have many text fragments that differ
    only in minor ways. This way you can list a single -OldText value
    (or a few) and have many -OldSubValue,-NewSubValue pairs to drive
    the replacement process.

    Each instance of -OldText will have its {{SubToken}} replaced with
    each -OldSubValue. Likewise, each instance of -NewText will be
    replaced with a copy of the original -OldText with {{SubToken}}
    replaced with the corresponding -NewSubValue.

    Finally, the modified pairs of -OldText,-NewText values will be used
    to replace text within your files (see -Files above).

    Here's an example:

        -OldText=Old text {{SubToken}} to be replaced.
        -NewText=(This will be ignored.)
        -OldSubValue=abc
        -NewSubValue=def
        -OldSubValue=123
        -NewSubValue=456

        Here are the resulting -OldText,-NewText pairs that would be used:

        -OldText=Old text abc to be replaced.
        -NewText=Old text def to be replaced.
        -OldText=Old text 123 to be replaced.
        -NewText=Old text 456 to be replaced.

-UseRegularExpressions=False

    Set this switch True to use regular expressions in the -OldText values
    (see above). Even with -UseRegularExpressions=False, you can still use
    standard regular expression escapes to replace most special characters
    (be sure to set -UseSpecialCharacters=True, see below).

-UseSpecialCharacters=False

    Set this switch True to use special characters in the -OldText
    values (see above). Here are the supported special characters:

    \b  = bell
    \e  = escape
    \f  = formfeed
    \n  = linefeed
    \q  = double quote
    \r  = carriage return
    \s  = single quote
    \t  = tab
    \v  = vertical tab

-XML_Profile=False

    Set this switch True to change the profile file from command-line format
    to XML format.


Notes:

    There may be various other settings that can be adjusted also (user
    interface settings, etc). See the profile file (""{INI}"")
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added ""on the fly""
    (in order of execution) to ""{INI}"" as the software runs.

"
                        .Replace("lsSubReplacementToken", lsSubReplacementToken)
                        .Replace("mcsFoundTextKey", mcsFoundTextKey)
                        .Replace("mcsFileChangedKey", mcsFileChangedKey)
                        .Replace("mcsFileUnchangedKey", mcsFileUnchangedKey)
                        .Replace("{EXE}", System.AppDomain.CurrentDomain.FriendlyName)
                        .Replace("{INI}", Path.GetFileName(loProfile.sActualPathFile))
                        .Replace("{{", "{")
                        .Replace("}}", "}")
                        );

                if ( lbUseRegularExpressions )
                {
                    leRegexOptions = (RegexOptions)loProfile.iValue("-RegexOptions", (int)RegexOptions.Singleline);
                    loProfile.GetAdd("-RegexOptionsHelp",
                            @"
Apply a boolean ""or"" to any of the following:

    RegexOptions.ExplicitCapture = 4
    RegexOptions.IgnoreCase = 1
    RegexOptions.IgnorePatternWhitespace = 32
    RegexOptions.Multiline = 2
    RegexOptions.None = 0
    RegexOptions.RightToLeft = 64
    RegexOptions.Singleline = 16
                            "
                            );
                }
                            lbNoPrompts = loProfile.bValue("-NoPrompts", false);
                            lbDisplayResults = loProfile.bValue("-DisplayResults", true);
                            lsDisplayResultsCommand = loProfile.sValue("-DisplayResultsModule", "Notepad.exe");
                bool        lbIgnoreNoFilesFound = loProfile.bValue("-IgnoreNoFilesFound", false);
                            lbCopyResultsToSTDOUT = loProfile.bValue("-CopyResultsToSTDOUT", false);
                bool        lbSearchOnly = loProfile.bValue("-SearchOnly", true);
                bool        lbRecurseSubdirectories = loProfile.bValue("-RecurseSubdirectories", true);
                bool        lbCaseInsensitiveSearch = loProfile.bValue("-CaseInsensitiveSearch", false);

                            // Fetch source code.
                            if ( loProfile.bValue("-FetchSource", false) )
                                tvFetchResource.ToDisk(System.Windows.Application.ResourceAssembly.GetName().Name, "ReplaceText.zip", null);

                string[]    lsFilesToReplacePathFilesArray = loProfile.sOneKeyArray("-Files");
                            loProfile.GetAdd("-Files", "One of many 'files to replace' pathfile(s) specifications goes here.");
                tvProfile   loOldTextToReplaceProfile = loProfile.oOneKeyProfile("-OldText", false);
                            loProfile.GetAdd("-OldText", "One of many old text substrings to replace goes here.");
                            this.ReplaceSpecialCharacters(loOldTextToReplaceProfile);
                string[]    lsOldTextToReplaceArray = loOldTextToReplaceProfile.sOneKeyArrayNoTrim("-OldText");
                string[]    lsOldTextToReplaceArrayToLower = null;
                string[]    lsOldTextToReplaceArrayToUpper = null;
                            if ( lbCaseInsensitiveSearch )
                            {
                                lsOldTextToReplaceArrayToLower = new string[lsOldTextToReplaceArray.Length];
                                lsOldTextToReplaceArrayToUpper = new string[lsOldTextToReplaceArray.Length];

                                for (int i=0; i < lsOldTextToReplaceArray.Length; i++)
                                {
                                    lsOldTextToReplaceArrayToLower[i] = lsOldTextToReplaceArray[i].ToLower();
                                    lsOldTextToReplaceArrayToUpper[i] = lsOldTextToReplaceArray[i].ToUpper();
                                }
                            }
                string[]    lsNewTextArray = loProfile.sOneKeyArrayNoTrim("-NewText");
                            loProfile.GetAdd("-NewText", "One of many new text replacement substrings goes here.");
                            this.ReplaceSpecialCharacters(lsNewTextArray);

                            if ( lsNewTextArray.Length != lsOldTextToReplaceArray.Length )
                                lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine)
                                        + "The number of -NewText items must match the number of -OldText items.";

                tvProfile   loProfileFilesChanged = new tvProfile();
                tvProfile   loProfileFilesUnChanged = new tvProfile();
                tvProfile   loProfileFilesFoundText = new tvProfile();
                bool        lbFilesFound = false;
                bool        lbErrorOtherThanFilesNotFound = false;

                if ( null == lsErrorMessages )
                {
                    foreach (string lsFilesToReplacePathFiles in lsFilesToReplacePathFilesArray)
                    {
                        System.Windows.Forms.Application.DoEvents();

                        bool lbLoopFilesFound = false;
                        bool lbLoopErrorOtherThanFilesNotFound = false;

                        string lsErrors = this.ReplacePathFiles(
                                  lsFilesToReplacePathFiles
                                , ref lsOldSubReplacementArray
                                , ref lsNewSubReplacementArray
                                , ref loOldTextToReplaceProfile
                                , ref lsOldTextToReplaceArray
                                , ref lsOldTextToReplaceArrayToLower
                                , ref lsOldTextToReplaceArrayToUpper
                                , ref lsNewTextArray
                                , ref loProfileFilesChanged
                                , ref loProfileFilesUnChanged
                                , ref loProfileFilesFoundText
                                , ref loProfile
                                , out lbLoopFilesFound
                                , out lbLoopErrorOtherThanFilesNotFound
                                );

                        if ( null != lsErrors)
                            lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine) + lsErrors;
                        if ( !lbFilesFound )
                            lbFilesFound = lbLoopFilesFound;
                        if ( !lbErrorOtherThanFilesNotFound )
                            lbErrorOtherThanFilesNotFound = lbLoopErrorOtherThanFilesNotFound;
                    }
                }

                loProfile.Remove(mcsFileChangedKey);
                loProfile.Remove(mcsFoundTextKey);
                loProfile.Remove(mcsFileUnchangedKey);

                loProfile.LoadFromCommandLineArray(loProfileFilesChanged.sCommandLineArray(), tvProfileLoadActions.Append);
                loProfile.LoadFromCommandLineArray(loProfileFilesFoundText.sCommandLineArray(), tvProfileLoadActions.Append);
                loProfile.LoadFromCommandLineArray(loProfileFilesUnChanged.sCommandLineArray(), tvProfileLoadActions.Append);
                loProfile.Save();

                if ( lbCopyResultsToSTDOUT )
                {
                    AttachConsole(ATTACH_PARENT_PROCESS);

                    if ( 0 != loProfileFilesChanged.Count )
                        Console.Write(loProfileFilesChanged.sCommandBlock());
                    if ( 0 != loProfileFilesFoundText.Count )
                        Console.Write(loProfileFilesFoundText.sCommandBlock());
                }

                if ( !lbFilesFound && !lbIgnoreNoFilesFound )
                    lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine)
                            + "No files could be found!";
            }
            catch (Exception ex)
            {
                lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine) + ex.Message;
            }

            if ( null != lsErrorMessages )
            {
                Environment.ExitCode = 1;

                if ( !lbNoPrompts )
                    System.Windows.Forms.MessageBox.Show(lsErrorMessages, this.Text);

                if ( lbCopyResultsToSTDOUT )
                    Console.Write(lsErrorMessages);
            }

            string[]    lsKeysArray = new string[]{mcsFileChangedKey, mcsFileUnchangedKey, mcsFoundTextKey};
            tvProfile   loOrderedProfile = new tvProfile();
                        foreach(DictionaryEntry loEntry in loProfile)
                            if ( lsKeysArray.Any(loEntry.Key.ToString().Contains) )
                                loOrderedProfile.Add(loEntry);

                        foreach(DictionaryEntry loEntry in loProfile)
                            if ( !lsKeysArray.Any(loEntry.Key.ToString().Contains) )
                                loOrderedProfile.Add(loEntry);

                        loProfile.Remove("*");

                        foreach(DictionaryEntry loEntry in loOrderedProfile)
                            loProfile.Add(loEntry);

                        loProfile.Save();

            this.Close();

            if ( lbDisplayResults )
            {
                loProfile.bEnableFileLock = false;
                Process.Start(lsDisplayResultsCommand, loProfile.sLoadedPathFile);
            }
        }

        private string ReplacePathFiles(
              string asPathFiles
            , ref string[] asOldSubReplacementArray
            , ref string[] asNewSubReplacementArray
            , ref tvProfile aoOldTextToReplaceProfile
            , ref string[] asOldTextToReplaceArray
            , ref string[] asOldTextToReplaceArrayToLower
            , ref string[] asOldTextToReplaceArrayToUpper
            , ref string[] asNewTextArray
            , ref tvProfile aoProfileFilesChanged
            , ref tvProfile aoProfileFilesUnChanged
            , ref tvProfile aoProfileFilesFoundText
            , ref tvProfile aoProfile
            , out bool abFilesFound
            , out bool abErrorOtherThanFilesNotFound
            )
        {
            string      lsErrorMessages = null;
            string      lsSubReplacementToken = aoProfile.sValue("-SubToken", "{SubToken}");
            bool        lbNoPrompts = aoProfile.bValue("-NoPrompts", false);
            bool        lbIgnoreNoFilesFound = aoProfile.bValue("-IgnoreNoFilesFound", false);
            bool        lbSearchOnly = aoProfile.bValue("-SearchOnly", false);
            bool        lbCaseInsensitiveSearch = aoProfile.bValue("-CaseInsensitiveSearch", false);
            bool        lbUseRegularExpressions = aoProfile.bValue("-UseRegularExpressions", false);
            RegexOptions leRegexOptions = RegexOptions.None;
                        if ( lbUseRegularExpressions )
                            leRegexOptions = (RegexOptions)aoProfile.iValue("-RegexOptions", (int)RegexOptions.Singleline);
            bool        lbRecurseSubdirectories = aoProfile.bValue("-RecurseSubdirectories", false);
            bool        lbHasToken = false;
                        if ( !String.IsNullOrEmpty(lsSubReplacementToken) )
                            for (int i=0; i < asOldTextToReplaceArray.Length; i++)
                            {
                                lbHasToken = asOldTextToReplaceArray[i].IndexOf(lsSubReplacementToken) > -1;
                                if ( lbHasToken )
                                    break;
                            }
            string      lsFilesToReplacePath = Path.GetDirectoryName(asPathFiles);
                        if ( "" == lsFilesToReplacePath )
                            lsFilesToReplacePath = ".";
            string      lsFilesToReplaceFiles = Path.GetFileName(asPathFiles);
            string[]    lsFilesToReplacePathFilesArray = null;
                        if ( Directory.Exists(lsFilesToReplacePath) )
                            lsFilesToReplacePathFilesArray = Directory.GetFiles(lsFilesToReplacePath, lsFilesToReplaceFiles);

            abFilesFound = false;
            abErrorOtherThanFilesNotFound = false;

            if ( null != lsFilesToReplacePathFilesArray && 0 != lsFilesToReplacePathFilesArray.Length )
            {
                abFilesFound = true;

                foreach (string lsFilesToReplacePathFile in lsFilesToReplacePathFilesArray)
                {
                    // Don't include the profile file currently in use.
                    if ( Path.GetFullPath(lsFilesToReplacePathFile) == msCurrentProfileAbsPathFile )
                        continue;

                    System.Windows.Forms.Application.DoEvents();

                    string lsFilename = Path.GetFileName(lsFilesToReplacePathFile);

                    if ( !lbNoPrompts )
                    {
                        this.Activate();
                        lblMessage.Text = String.Format("{0} text in file \"{1}\" ...", lbSearchOnly ? "Searching" : "Replacing", lsFilename);
                        this.WindowState = FormWindowState.Normal;
                        this.Show();
                        this.Refresh();
                    }

                    bool lbHasOldText = false;

                    try
                    {
                        // Append a trailing newline character to allow for EOF block matches.
                        string lsFileAsStream = File.ReadAllText(lsFilesToReplacePathFile) + Environment.NewLine;
                        string lsOriginalFileAsStream = lsFileAsStream;
                        for (int liSubReplaceIndex = 0; liSubReplaceIndex < asOldSubReplacementArray.Length; liSubReplaceIndex++)
                        {
                            if ( lbHasToken )
                            {
                                // Reset the old text array before replacing tokens.
                                asOldTextToReplaceArray = aoOldTextToReplaceProfile.sOneKeyArrayNoTrim("-OldText");

                                for (int i=0; i < asOldTextToReplaceArray.Length; i++)
                                    asOldTextToReplaceArray[i] = asOldTextToReplaceArray[i]
                                            .Replace(lsSubReplacementToken, asOldSubReplacementArray[liSubReplaceIndex]);

                                // Redefine the new text array to that of the old before replacing tokens.
                                asNewTextArray = aoOldTextToReplaceProfile.sOneKeyArrayNoTrim("-OldText");

                                for (int i=0; i < asNewTextArray.Length; i++)
                                    asNewTextArray[i] = asNewTextArray[i]
                                            .Replace(lsSubReplacementToken, asNewSubReplacementArray[liSubReplaceIndex]);
                            }

                            if ( !lbCaseInsensitiveSearch )
                            {
                                if ( lbUseRegularExpressions )
                                {
                                    // Replace old with new text substrings.
                                    for (int i=0; i < asOldTextToReplaceArray.Length; i++)
                                    {
                                        Regex loRegex = new Regex(asOldTextToReplaceArray[i], leRegexOptions);

                                        if (!lbHasOldText && loRegex.IsMatch(lsFileAsStream))
                                            lbHasOldText = true;

                                        lsFileAsStream = loRegex.Replace(lsFileAsStream, asNewTextArray[i]);
                                    }
                                }
                                else
                                {
                                    StringBuilder lsbFileAsStream = new StringBuilder(lsFileAsStream);

                                    // Replace old with new text substrings.
                                    for (int i=0; i < asOldTextToReplaceArray.Length; i++)
                                    {
                                        if ( !lbHasOldText && lsOriginalFileAsStream.IndexOf(asOldTextToReplaceArray[i]) > -1 )
                                            lbHasOldText = true;

                                        lsbFileAsStream.Replace(asOldTextToReplaceArray[i], asNewTextArray[i]);
                                    }

                                    lsFileAsStream = lsbFileAsStream.ToString();
                                }
                            }
                            else
                            {
                                if ( lbUseRegularExpressions )
                                {
                                    // Replace old with new text substrings.
                                    for (int i=0; i < asOldTextToReplaceArray.Length; i++)
                                    {
                                        Regex loRegex = new Regex(asOldTextToReplaceArray[i], leRegexOptions);
                                        Regex loToLowerRegex = new Regex(asOldTextToReplaceArrayToLower[i], leRegexOptions);
                                        Regex loToUpperRegex = new Regex(asOldTextToReplaceArrayToUpper[i], leRegexOptions);

                                        if ( !lbHasOldText && loRegex.IsMatch(lsFileAsStream) )
                                            lbHasOldText = true;
                                        if ( !lbHasOldText && loToLowerRegex.IsMatch(lsFileAsStream) )
                                            lbHasOldText = true;
                                        if ( !lbHasOldText && loToUpperRegex.IsMatch(lsFileAsStream) )
                                            lbHasOldText = true;

                                        lsFileAsStream = loRegex.Replace(lsFileAsStream, asNewTextArray[i]);
                                        lsFileAsStream = loToLowerRegex.Replace(lsFileAsStream, asNewTextArray[i]);
                                        lsFileAsStream = loToUpperRegex.Replace(lsFileAsStream, asNewTextArray[i]);
                                    }
                                }
                                else
                                {
                                    StringBuilder lsbFileAsStream = new StringBuilder(lsFileAsStream);

                                    // Replace old with new text substrings.
                                    for (int i=0; i < asOldTextToReplaceArray.Length; i++)
                                    {
                                        if ( !lbHasOldText && lsOriginalFileAsStream.IndexOf(asOldTextToReplaceArray[i]) > -1 )
                                            lbHasOldText = true;
                                        if ( !lbHasOldText && lsOriginalFileAsStream.IndexOf(asOldTextToReplaceArrayToLower[i]) > -1 )
                                            lbHasOldText = true;
                                        if ( !lbHasOldText && lsOriginalFileAsStream.IndexOf(asOldTextToReplaceArrayToUpper[i]) > -1 )
                                            lbHasOldText = true;

                                        lsbFileAsStream.Replace(asOldTextToReplaceArray[i], asNewTextArray[i]);
                                        lsbFileAsStream.Replace(asOldTextToReplaceArrayToLower[i], asNewTextArray[i]);
                                        lsbFileAsStream.Replace(asOldTextToReplaceArrayToUpper[i], asNewTextArray[i]);
                                    }

                                    lsFileAsStream = lsbFileAsStream.ToString();
                                }
                            }
                        }

                        if ( lbSearchOnly || lsFileAsStream == lsOriginalFileAsStream )
                        {
                            if ( lbHasOldText )
                                aoProfileFilesFoundText.Add(mcsFoundTextKey, lsFilesToReplacePathFile);
                            else
                            if ( !lbSearchOnly )
                                aoProfileFilesUnChanged.Add(mcsFileUnchangedKey, lsFilesToReplacePathFile);
                        }
                        else
                        {
                            StreamWriter loStreamWriter = new StreamWriter(lsFilesToReplacePathFile, false);

                            // Remove the extra trailing newline (added support for EOF block matches, see above).
                            loStreamWriter.Write(lsFileAsStream.Substring(0, lsFileAsStream.Length - Environment.NewLine.Length));
                            loStreamWriter.Close();

                            aoProfileFilesChanged.Add(mcsFileChangedKey, lsFilesToReplacePathFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        abErrorOtherThanFilesNotFound = true;
                        lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine)
                                + String.Format(
                                  "Error: \"{0}\" occurred while attempting to write to the file: \"{1}\"."
                                , ex.Message
                                , lsFilename
                                );
                    }
                } // End files loop
            }

            if ( lbRecurseSubdirectories )
            {
                // Process the sub-folders in the base folder.

                // Use an empty array instead of null to
                // prevent the "foreach" from blowing up.
                string[]    lsSubfoldersArray = new string[0];
                            if ( Directory.Exists(lsFilesToReplacePath) )
                            {
                                try
                                {
                                    // Get subdirectories only at the next level.
                                    lsSubfoldersArray = Directory.GetDirectories(lsFilesToReplacePath);
                                }
                                catch (Exception ex)
                                {
                                    lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine) + ex.Message;
                                }
                            }

                foreach (string lsSubfolder in lsSubfoldersArray)
                {
                    System.Windows.Forms.Application.DoEvents();

                    bool lbLoopFilesFound = false;
                    bool lbLoopErrorOtherThanFilesNotFound = false;

                    // Replace text in all applicable files in the current subfolder.
                    string lsErrors = this.ReplacePathFiles(
                              Path.Combine(lsSubfolder, lsFilesToReplaceFiles)
                            , ref asOldSubReplacementArray
                            , ref asNewSubReplacementArray
                            , ref aoOldTextToReplaceProfile
                            , ref asOldTextToReplaceArray
                            , ref asOldTextToReplaceArrayToLower
                            , ref asOldTextToReplaceArrayToUpper
                            , ref asNewTextArray
                            , ref aoProfileFilesChanged
                            , ref aoProfileFilesUnChanged
                            , ref aoProfileFilesFoundText
                            , ref aoProfile
                            , out lbLoopFilesFound
                            , out lbLoopErrorOtherThanFilesNotFound
                            );

                    if ( null != lsErrors)
                        lsErrorMessages += (null == lsErrorMessages ? "" : Environment.NewLine + Environment.NewLine) + lsErrors;
                    if ( !abFilesFound )
                        abFilesFound = lbLoopFilesFound;
                    if ( !abErrorOtherThanFilesNotFound )
                        abErrorOtherThanFilesNotFound = lbLoopErrorOtherThanFilesNotFound;
                }
            }

            return lsErrorMessages;
        }

        private void ReplaceSpecialCharacters(IList aoList)
        {
            const char mccHideDoubleBackslash = '\u0001';

            for (int i=0; i < aoList.Count; i++)
            {
                aoList[i] = ((string)aoList[i]).Replace("\\\\", mccHideDoubleBackslash.ToString());
            }

            if ( mbUseSpecialCharacters && !mbSpecialCharactersProfileInitDone )
            {
                int liCount = moSpecialCharactersProfile.Count;

                for (int i=0; i<liCount; i++)
                    moSpecialCharactersProfile.Add(moSpecialCharactersProfile.oEntry(i).Key.ToString().ToUpper(), moSpecialCharactersProfile.oEntry(i).Value);
                
                mbSpecialCharactersProfileInitDone = true;
            }

            if ( mbUseSpecialCharacters )
                for (int i=0; i < aoList.Count; i++)
                {
                    foreach (DictionaryEntry loEntry in moSpecialCharactersProfile)
                        aoList[i] = ((string)aoList[i]).Replace(loEntry.Key.ToString(), loEntry.Value.ToString());
                }

            for (int i=0; i < aoList.Count; i++)
            {
                // Replace each double backslash with a single backslash.
                aoList[i] = ((string)aoList[i]).Replace(mccHideDoubleBackslash.ToString(), "\\");
            }
        }
    }
}
