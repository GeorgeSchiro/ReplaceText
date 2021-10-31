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

-Files="One of many 'files to replace' pathfile(s) specifications goes here."

    This is the specification of files to be processed. It can include
    wildcards.

    Note: This key may appear any number of times in the profile.

-Help= SEE PROFILE FOR DEFAULT VALUE

    This help text.

-IgnoreNoFilesFound=False

    Set this switch True and no error pop-up will appear if no files are
    actually found to process.

-NewSubValue="One of many new 'sub replacement' values goes here."

    This is a new "sub replacement" value to replace the corresponding
    -OldSubValue (see below) within each -OldText value (see below).

    See -SubToken below for more details.

    Note: This key may appear any number of times in the profile.

-NewText="One of many new text replacement substrings goes here."

    This is a new text substring to replace a corresponding old text
    substring (see -OldText below) in all of the files given by the
    various -Files specifications (see above).

    Note: This key may appear any number of times in the profile.

-NoPrompts=False

    Set this switch True and all pop-up prompts will be suppressed. You must
    use this switch whenever the software is run via a server computer batch
    job or job scheduler (ie. where no user interaction is permitted).

-OldSubValue="One of many old 'sub replacement' values goes here."

    This is an old "sub replacement" value to replace the corresponding
    -NewSubValue (see above) within each -OldText value (see below).

    See -SubToken below for more details.

    Note: This key may appear any number of times in the profile.

-OldText="One of many old text substrings to replace goes here."

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
    information will be written to the profile file (ie. "ReplaceText.exe.txt").

-SearchOnly=True

    Set this switch False and all files matching the specifications in
    the -Files parameters (see above) will be updated if they contain at
    least one -OldText value (see above). Otherwise, each matching file
    that contains at least one -OldText value will be displayed with the
    -FoundIn key.

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the
    profile file at startup in command-line format. This may be helpful as a
    diagnostic.

-SubToken="{SubToken}"

    A "sub replacement token" ({SubToken}) can be used to
    pass a common substring value referenced as (-OldSubValue) to be
    replaced with the corresponding -NewSubValue (see above) within each
    (-OldText) value.

    Any number of -OldSubValue,-NewSubValue pairs can be given, all of
    which will be replaced in every -OldText value (if found there).

    This feature is useful if you have many text fragments that differ
    only in minor ways. This way you can list a single -OldText value
    (or a few) and have many -OldSubValue,-NewSubValue pairs to drive
    the replacement process.

    Each instance of -OldText will have its {SubToken} replaced with
    each -OldSubValue. Likewise, each instance of -NewText will be
    replaced with a copy of the original -OldText with {SubToken}
    replaced with the corresponding -NewSubValue.

    Finally, the modified pairs of -OldText,-NewText values will be used
    to replace text within your files (see -Files above).

    Here's an example:

        -OldText=Old text {SubToken} to be replaced.
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
    interface settings, etc). See the profile file ("ReplaceText.exe.txt")
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added "on the fly"
    (in order of execution) to "ReplaceText.exe.txt" as the software runs.
