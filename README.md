Overview
========


<b>ReplaceText</b> is a simple mass search & replace
tool for many text files in many locations.

This utility will take any files (presumably text) referenced as -Files and
replace old text referenced as -OldText with new text referenced as -NewText.


Features
========


-   Simple setup - try it out fast
-   Accepts wildcards in file specifications
-   Has a "search only" option (the default)
-   Supports regular expressions
-   Can impact an entire subdirectory tree
-   Can impact any number of disk volumes
-   Can be command-line driven from another process
-   Software is highly configurable
-   Software is totally self-contained (EXE is its own setup)


Details
=======


<b>ReplaceText</b> was designed for programmers.

It is a greatly simplified way of doing what you might otherwise
do with far more complex *nix tools like perl, grep or sed.

<b>ReplaceText</b> is optimized for simple mass text manipulation tasks.

All of its flexibility is managed through a single plain-text profile file
(ie. a configuration file). Everything is managed through the profile file.
<b>ReplaceText</b> does not use the windows registry at all.

You will notice the profile data (what drives the search & replace process)
is not formatted as XML. It is expressed in simple "command-line" format.
That makes it easier to read and parse. It is also a reminder that anything you 
see in the profile file can be overridden with the equivalent "-key=value" pairs
passed on the "ReplaceText.exe" command-line.


Requirements
============


-   .Net Framework 3.5+


Command-Line Usage
==================


    Open this utility's profile file to see additional options available. It
    is usually located in the same folder as "ReplaceText.exe" and has the same
    name with ".txt" added (see "ReplaceText.exe.txt").

    Profile file options can be overridden with command-line arguments. The
    keys for any "-key=value" pairs passed on the command-line must match
    those that appear in the profile (with the exception of the "-ini" key).

    For example, the following invokes the use of an alternative profile file:

        ReplaceText.exe -ini=NewProfile.txt

    This tells the software to display all files changed and unchanged:

        ReplaceText.exe -DisplayResults


    Author:  George Schiro (GeoCode@Schiro.name)

    Date:    1/27/2005

 
Options and Features
====================


    The main options for this utility are listed below with their default values.
    A brief description of each feature follows.

-CaseInsensitiveSearch=False

    Set this switch True and the original version plus both uppercase and
    lowercase versions of each given -OldText string (not mixed-case versions)
    will be replaced (see -OldText below).

-CopyResultsToSTDOUT=False

    Set this switch True and replacement results will be copied to standard
    output. This can be useful if the software is run by another process.

-DisplayResults=True

    Set this switch False and replacement results will not be displayed using
    the -DisplayResultsModule (see below).

-DisplayResultsModule="Notepad.exe"

    This is the software used to display the replacement results.

    Note: the standard filetype association will be used if it's empty.

-FetchSource=False

    Set this switch True to fetch the source code for this utility
    from the EXE. Look in the containing folder for a ZIP file with
    the full project sources.

-Files="One of many 'files to replace' pathfile(s) specifications goes here."

    This is the specification of files to be processed. It can include
    wildcards.

    Note: This key may appear any number of times in the profile.

-Help= SEE PROFILE FOR DEFAULT VALUE

    This help text.

-IgnoreIOException=False

    Set this switch True to ignore any IOException during file reads.

-IgnoreNoFilesFound=False

    Set this switch True to ignore "no files found" errors.

-IgnoreUnauthorizedAccess=False

    Set this switch True to ignore any UnauthorizedAccessException
    during file or directory reads.

-ListSubTokenReplacements=False

    The list of -SubToken values together with -OldSubValue, -NewSubValue
    pairs are used to modify the -OldText, -NewText pairs (see below).

    Set this switch True to have the modified -OldText, -NewText pairs
    added to the profile for your perusal after processing completes.

-NewSubValue="One of many new 'sub replacement' values goes here."

    This is a new "sub replacement" value used in lieu of the corresponding
    -OldSubValue (see below) in each -NewText (or, if -NewText is empty, a
    copy of the corresponding -OldText) to replace an embedded -SubToken.
    If -NewSubValue is empty, the corresponding -OldSubValue will be used.

    See -SubToken below for more details.

    Note: This key may appear any number of times in the profile.

-NewText="One of many new text replacement strings goes here."

    This is a new text string to replace a corresponding old text string
    (see -OldText below) in all of the files given by the various -Files
    specifications (see above).

    Note: This key may appear any number of times in the profile.

-NoPrompts=False

    Set this switch True and all pop-up prompts will be suppressed. You must
    use this switch whenever the software is run via a server computer batch
    job or job scheduler (ie. where no user interaction is permitted).

-OldSubValue="One of many old 'sub replacement' values goes here."

    This is an old "sub replacement" value used to replace an embedded
    -SubToken within each -OldText string (see below).

    See -SubToken below for more details.

    Note: This key may appear any number of times in the profile.

-OldText="One of many old text strings to replace goes here."

    This is an old text string to be replaced by a corresponding new text
    string (see -NewText above) in all of the files given by the various
    -Files specifications (see above).

    Note: This key may appear any number of times in the profile.

-RecurseSubdirectories=True

    Set this switch False and only the base directory found in each
    -Files specification (see above) will be searched for -OldText
    (see above). Otherwise, every file matching the -Files specifications
    found in every subdirectory from each base subdirectory onward will
    be searched.

    Note: BE CAREFUL!!! Hundreds or thousands of files could be impacted!

    Always use "-SearchOnly=True" for the first run. That way you can
    see the scope of any potential damage before the damage is done.

-SaveProfile=True

    Set this switch False to prevent saving to the profile file by this
    software. This is not recommended since status information is written
    to the profile after each run.

-SaveSansCmdLine=True

    Set this switch False to allow merged command-lines to be written to
    the profile file (ie. "ReplaceText.exe.txt"). When True, everything
    but command-line keys will be saved.

-SearchOnly=True

    Set this switch False and all files matching the specifications in
    the -Files parameters (see above) will be updated if they contain at
    least one -OldText string (see above). Otherwise, each matching file
    that contains at least one -OldText string will be displayed with the
    -FoundIn key.

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the
    profile file at startup in command-line format. This may be helpful as
    a diagnostic.

-SubToken="One of many 'sub replacement' tokens goes here."

    A "sub replacement token" can be used to pass common substring values
    (referenced as -OldSubValue, see above) into various -OldText strings.
    The same token can also be used to pass separate common substring values
    (referenced as -NewSubValue, see above) into various -NewText strings.

    Any number of -OldSubValue, -NewSubValue pairs can be inserted into
    various -OldText, -NewText pairs that contain at least one -SubToken.

    If there are fewer -SubToken values than -OldSubValue, -NewSubValue
    pairs, the last -SubToken defined will be used for the balance of
    -OldSubValue, -NewSubValue pairs.

    This feature is useful if you have many text fragments that differ
    only in minor ways. This way you can list a single -OldText string
    (or a few) and have many -OldSubValue, -NewSubValue pairs driving the
    replacement process with the various sub-replacements filled-in.

    Suppose -SubToken="{SubToken}". Each instance of -OldText will have
    its {SubToken} replaced with each -OldSubValue. Likewise, each instance
    of -NewText will have its {SubToken} replaced with each -NewSubValue.
    If -NewText is empty, it will be replaced with a copy of the original
    -OldText with {SubToken} replaced with each -NewSubValue.

    Finally, the modified list of -OldText, -NewText pairs will then be
    used to replace text within your files (see -Files above).

    Here's an example:

        -SubToken={ST1}
        -OldSubValue=abc
        -NewSubValue=def
        -SubToken={ST2}
        -OldSubValue=123
        -NewSubValue=456
        -OldSubValue=uvw
        -NewSubValue=xyz
        -OldText=Old text {ST1} to be replaced.
        -NewText=
        -OldText=More old text {ST2} to be replaced.
        -NewText=New text {ST2} now in its place.

        Here are the -OldText, -NewText pairs that would result:

        -OldText=Old text abc to be replaced.
        -NewText=Old text def to be replaced.
        -OldText=More old text 123 to be replaced.
        -NewText=New text 456 now in its place.
        -OldText=More old text uvw to be replaced.
        -NewText=New text xyz now in its place.

-TrackItemsFoundPerFile=False

    Set this switch True to have every -OldText item found in every file
    tracked during the current run. Each -OldText item found is compared
    to corresponding items found during the previous run. If there are
    any discrepancies found between the current run and the previous run,
    a warning dialog will be displayed (assuming -NoPrompts is false, see
    above). This can be helpful to catch inadvertent manual text edits that
    often break software otherwise maintained by automated code transforms
    implemented via a tool like this.

-UseRegularExpressions=False

    Set this switch True to use regular expressions in -OldText strings
    (see above) as well as capture groups in the corresponding -NewText
    strings. Even with -UseRegularExpressions set False, you can still use
    standard regular expression escapes to replace most special characters
    (be sure to set -UseSpecialCharacters=True, see below).

-UseSpecialCharacters=False

    Set this switch True to use special characters in -OldText strings
    as well as in -NewText strings (see above).

    Here are the supported special characters:

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
    interface settings, etc). See the profile file ("ReplaceText.exe.txt")
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added "on the fly"
    (in order of execution) to "ReplaceText.exe.txt" as the software runs.
