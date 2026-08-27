# CLI Command Reference

> Auto-generated recursively from the built `pptcli` runtime. Do not edit by hand.

## `pptcli accessibility`

```text
DESCRIPTION:
Deterministic presentation accessibility checks and reading-order operations

USAGE:
    pptcli accessibility <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                            Prints help information
    -s, --session <SESSION>               Session ID from 'session open' command
        --slide-index <SLIDEINDEX>        (required for: get-reading-order,
                                          set-reading-order) (valid for:
                                          get-reading-order, set-reading-order)
        --shape-indexes <SHAPEINDEXES>    (required for: set-reading-order)
                                          (valid for: set-reading-order) (JSON
                                          format)
    -o, --output <PATH>                   Write output to file instead of
                                          stdout. For image results, decodes and
                                          saves as binary file
```

## `pptcli animation`

```text
DESCRIPTION:
Animation commands: add/delete entrance, emphasis, and exit effects on a shape's
slide timeline (Slide.TimeLine.MainSequence), and read/set a slide's transition
(Slide.SlideShowTransition). Operates within an already-open , targeting a
specific slide (and, for shape effects, a specific shape) by 1-based index

USAGE:
    pptcli animation <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                         Prints help information
    -s, --session <SESSION>                            Session ID from 'session
                                                       open' command
        --slide-index <SLIDEINDEX>                     (required)
        --shape-index <SHAPEINDEX>                     (required for:
                                                       add-effect) (valid for:
                                                       add-effect)
        --effect-name <EFFECTNAME>                     (required for:
                                                       add-effect) (valid for:
                                                       add-effect)
        --is-exit <ISEXIT>                             When true, the effect is
                                                       applied as the shape
                                                       leaving the slide (exit)
                                                       rather than the
                                                       default entrance/emphasis
                                                       behavior. (valid for:
                                                       add-effect)
        --trigger <TRIGGER>                            When the effect starts:
                                                       "on-click" (default),
                                                       "with-previous", or
                                                       "after-previous". (valid
                                                       for: add-effect)
        --effect-index <EFFECTINDEX>                   (required for:
                                                       delete-effect) (valid
                                                       for: delete-effect)
        --transition-name <TRANSITIONNAME>             (required for:
                                                       set-transition) (valid
                                                       for: set-transition)
        --duration-seconds <DURATIONSECONDS>           (valid for:
                                                       set-transition)
        --advance-on-click <ADVANCEONCLICK>            (valid for:
                                                       set-transition)
        --advance-on-time <ADVANCEONTIME>              (valid for:
                                                       set-transition)
        --advance-time-seconds <ADVANCETIMESECONDS>    (valid for:
                                                       set-transition)
    -o, --output <PATH>                                Write output to file
                                                       instead of stdout. For
                                                       image results, decodes
                                                       and saves as binary file
```

## `pptcli chart`

```text
DESCRIPTION:
Chart lifecycle, data, and quick-formatting operations

USAGE:
    pptcli chart <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                            Prints help information
    -s, --session <SESSION>               Session ID from 'session open' command
        --slide-index <SLIDEINDEX>        1-based slide index. (required)
        --chart-type <CHARTTYPE>          Chart type: "bar", "line", or "pie".
                                          (required for: add-chart) (valid for:
                                          add-chart)
        --left <LEFT>                     Left position in points. (required
                                          for: add-chart) (valid for: add-chart)
        --top <TOP>                       Top position in points. (required for:
                                          add-chart) (valid for: add-chart)
        --width <WIDTH>                   Width in points. (required for:
                                          add-chart) (valid for: add-chart)
        --height <HEIGHT>                 Height in points. (required for:
                                          add-chart) (valid for: add-chart)
        --categories <CATEGORIES>         Category labels (x-axis / pie slice
                                          labels). (required for: add-chart,
                                          replace-chart-data) (valid for:
                                          add-chart, replace-chart-data) (JSON
                                          format)
        --series-name <SERIESNAME>        Name of the single data series.
                                          (required for: add-chart, add-series)
                                          (valid for: add-chart, add-series)
        --values <VALUES>                 Data values, one per category.
                                          (required for: add-chart, add-series)
                                          (valid for: add-chart, add-series)
                                          (JSON format)
        --shape-index <SHAPEINDEX>        (required for: get-chart-data,
                                          add-series, set-chart-title,
                                          get-chart-title, set-axis-title,
                                          get-axis-title, set-legend-visibility,
                                          get-legend-visibility,
                                          replace-chart-data, get-style,
                                          set-style, get-color-style,
                                          set-color-style, get-data-table,
                                          set-data-table) (valid for:
                                          get-chart-data, add-series,
                                          set-chart-title, get-chart-title,
                                          set-axis-title, get-axis-title,
                                          set-legend-visibility,
                                          get-legend-visibility,
                                          replace-chart-data, get-style,
                                          set-style, get-color-style,
                                          set-color-style, get-data-table,
                                          set-data-table)
        --title <TITLE>                   (required for: set-chart-title,
                                          set-axis-title) (valid for:
                                          set-chart-title, set-axis-title)
        --axis-type <AXISTYPE>            (required for: set-axis-title,
                                          get-axis-title) (valid for:
                                          set-axis-title, get-axis-title)
        --visible <VISIBLE>               True to show the chart element; false
                                          to hide it. (required for:
                                          set-legend-visibility, set-data-table)
                                          (valid for: set-legend-visibility,
                                          set-data-table)
        --series-names <SERIESNAMES>      (required for: replace-chart-data)
                                          (valid for: replace-chart-data) (JSON
                                          format)
        --series-values <SERIESVALUES>    (required for: replace-chart-data)
                                          (valid for: replace-chart-data) (JSON
                                          format)
        --style <STYLE>                   Built-in chart style number, verified
                                          against PowerPoint from 1 through 48.
                                          (required for: set-style) (valid for:
                                          set-style)
        --color-style <COLORSTYLE>        Built-in chart color style number,
                                          verified against PowerPoint from 1
                                          through 26. (required for:
                                          set-color-style) (valid for:
                                          set-color-style)
    -o, --output <PATH>                   Write output to file instead of
                                          stdout. For image results, decodes and
                                          saves as binary file
```

## `pptcli export`

```text
DESCRIPTION:
Export commands: render presentations to PDF or slides to raster image files.
Operates within an already-open

USAGE:
    pptcli export <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                  Prints help information
    -s, --session <SESSION>                     Session ID from 'session open'
                                                command
        --output-path <OUTPUTPATH>              Full path for the output .pdf
                                                file. (required for:
                                                export-to-pdf,
                                                export-slide-to-image) (valid
                                                for: export-to-pdf,
                                                export-slide-to-image)
        --overwrite <OVERWRITE>                 Whether an existing PDF may be
                                                replaced. Defaults to false.
                                                (valid for: export-to-pdf)
        --slide-index <SLIDEINDEX>              1-based index of the slide to
                                                export. (required for:
                                                export-slide-to-image) (valid
                                                for: export-slide-to-image)
        --format <FORMAT>                       PowerPoint filter name for the
                                                image format (e.g. "PNG", "JPG",
                                                "GIF").     Defaults to "PNG".
                                                (valid for:
                                                export-slide-to-image,
                                                export-all-slides-to-images)
        --width <WIDTH>                         Optional output width in pixels;
                                                0 or null uses PowerPoint's
                                                default. (valid for:
                                                export-slide-to-image)
        --height <HEIGHT>                       Optional output height in
                                                pixels; 0 or null uses
                                                PowerPoint's default. (valid
                                                for: export-slide-to-image)
        --output-directory <OUTPUTDIRECTORY>    Directory where slide images
                                                will be written. Created if it
                                                does not exist.     PowerPoint
                                                names the output files
                                                Slide1.{ext}, Slide2.{ext}, etc.
                                                (required for:
                                                export-all-slides-to-images)
                                                (valid for:
                                                export-all-slides-to-images)
```

## `pptcli image`

```text
DESCRIPTION:
Image commands: add a picture file to a slide. Operates within an already-open
IPresentationBatch, targeting a specific slide by its 1-based index

USAGE:
    pptcli image <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                     Prints help information
    -s, --session <SESSION>                        Session ID from 'session
                                                   open' command
        --slide-index <SLIDEINDEX>                 (required)
        --image-path <IMAGEPATH>                   (required for: add-picture)
                                                   (valid for: add-picture)
        --left <LEFT>                              (required for: add-picture)
                                                   (valid for: add-picture)
        --top <TOP>                                (required for: add-picture)
                                                   (valid for: add-picture)
        --width <WIDTH>                            (required for: add-picture)
                                                   (valid for: add-picture)
        --height <HEIGHT>                          (required for: add-picture)
                                                   (valid for: add-picture)
        --link-to-file <LINKTOFILE>                Whether the picture remains
                                                   linked to its source file.
                                                   Defaults to false. (valid
                                                   for: add-picture)
        --save-with-document <SAVEWITHDOCUMENT>    Whether PowerPoint stores
                                                   picture data in the
                                                   presentation. Defaults to
                                                   true. (valid for:
                                                   add-picture)
        --shape-index <SHAPEINDEX>                 (required for:
                                                   set-brightness-contrast,
                                                   get-brightness-contrast,
                                                   set-recolor, get-recolor,
                                                   set-crop, get-crop) (valid
                                                   for: set-brightness-contrast,
                                                   get-brightness-contrast,
                                                   set-recolor, get-recolor,
                                                   set-crop, get-crop)
        --brightness <BRIGHTNESS>                  (required for:
                                                   set-brightness-contrast)
                                                   (valid for:
                                                   set-brightness-contrast)
        --contrast <CONTRAST>                      (required for:
                                                   set-brightness-contrast)
                                                   (valid for:
                                                   set-brightness-contrast)
        --color-type <COLORTYPE>                   (required for: set-recolor)
                                                   (valid for: set-recolor)
        --crop-left <CROPLEFT>                     (required for: set-crop)
                                                   (valid for: set-crop)
        --crop-top <CROPTOP>                       (required for: set-crop)
                                                   (valid for: set-crop)
        --crop-right <CROPRIGHT>                   (required for: set-crop)
                                                   (valid for: set-crop)
        --crop-bottom <CROPBOTTOM>                 (required for: set-crop)
                                                   (valid for: set-crop)
    -o, --output <PATH>                            Write output to file instead
                                                   of stdout. For image results,
                                                   decodes and saves as binary
                                                   file
```

## `pptcli layout`

```text
DESCRIPTION:
Slide layout commands: apply/read a slide's built-in layout. Operates within an
already-open IPresentationBatch, targeting a specific slide by its 1-based index

USAGE:
    pptcli layout <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                          Prints help information
    -s, --session <SESSION>             Session ID from 'session open' command
        --slide-index <SLIDEINDEX>      (required for: set-layout, get-layout)
                                        (valid for: set-layout, get-layout)
        --layout-name <LAYOUTNAME>      (required for: set-layout) (valid for:
                                        set-layout)
        --master-index <MASTERINDEX>    (required for: list-layouts,
                                        delete-layout) (valid for: list-layouts,
                                        delete-layout)
        --layout-index <LAYOUTINDEX>    (required for: delete-layout) (valid
                                        for: delete-layout)
    -o, --output <PATH>                 Write output to file instead of stdout.
                                        For image results, decodes and saves as
                                        binary file
```

## `pptcli master`

```text
DESCRIPTION:
Slide master commands: read/edit the title and body placeholder fonts on the
presentation's slide master, and read/edit the slide master's background fill
color. Operates within an already-open . Changes here apply to every slide that
inherits from the master (i.e. any slide that does not itself override the
property), which is the practical "edit the master, not each slide" workflow
PowerPoint's COM object model supports safely

USAGE:
    pptcli master <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                  Prints help information
    -s, --session <SESSION>                     Session ID from 'session open'
                                                command
        --font-name <FONTNAME>                  (valid for: set-title-font,
                                                set-body-font)
        --font-size <FONTSIZE>                  (valid for: set-title-font,
                                                set-body-font)
        --bold <BOLD>                           (valid for: set-title-font,
                                                set-body-font)
        --red <RED>                             (required for:
                                                set-background-color) (valid
                                                for: set-title-font,
                                                set-body-font,
                                                set-background-color)
        --green <GREEN>                         (required for:
                                                set-background-color) (valid
                                                for: set-title-font,
                                                set-body-font,
                                                set-background-color)
        --blue <BLUE>                           (required for:
                                                set-background-color) (valid
                                                for: set-title-font,
                                                set-body-font,
                                                set-background-color)
        --red1 <RED1>                           (required for:
                                                set-gradient-background) (valid
                                                for: set-gradient-background)
        --green1 <GREEN1>                       (required for:
                                                set-gradient-background) (valid
                                                for: set-gradient-background)
        --blue1 <BLUE1>                         (required for:
                                                set-gradient-background) (valid
                                                for: set-gradient-background)
        --red2 <RED2>                           (required for:
                                                set-gradient-background) (valid
                                                for: set-gradient-background)
        --green2 <GREEN2>                       (required for:
                                                set-gradient-background) (valid
                                                for: set-gradient-background)
        --blue2 <BLUE2>                         (required for:
                                                set-gradient-background) (valid
                                                for: set-gradient-background)
        --gradient-style <GRADIENTSTYLE>        (valid for:
                                                set-gradient-background)
        --gradient-variant <GRADIENTVARIANT>    (valid for:
                                                set-gradient-background)
        --master-index <MASTERINDEX>            (required for: delete-master)
                                                (valid for: delete-master)
    -o, --output <PATH>                         Write output to file instead of
                                                stdout. For image results,
                                                decodes and saves as binary file
```

## `pptcli notes`

```text
DESCRIPTION:
Speaker notes commands: set/get the notes text for a slide. Operates within an
already-open IPresentationBatch, targeting a specific slide by its 1-based index

USAGE:
    pptcli notes <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                        Prints help information
    -s, --session <SESSION>           Session ID from 'session open' command
        --slide-index <SLIDEINDEX>    (required)
        --text <TEXT>                 (required for: set-notes-text) (valid for:
                                      set-notes-text)
    -o, --output <PATH>               Write output to file instead of stdout.
                                      For image results, decodes and saves as
                                      binary file
```

## `pptcli pagesetup`

```text
DESCRIPTION:
Presentation-wide slide size, numbering, and footer settings

USAGE:
    pptcli pagesetup <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                        Prints help information
    -s, --session <SESSION>                           Session ID from 'session
                                                      open' command
        --width <WIDTH>                               (required for: set-size)
                                                      (valid for: set-size)
        --height <HEIGHT>                             (required for: set-size)
                                                      (valid for: set-size)
        --first-slide-number <FIRSTSLIDENUMBER>       (required for:
                                                      set-first-slide-number)
                                                      (valid for:
                                                      set-first-slide-number)
        --footer-text <FOOTERTEXT>                    (valid for: set-footer)
        --show-footer <SHOWFOOTER>                    (valid for: set-footer)
        --show-slide-number <SHOWSLIDENUMBER>         (valid for: set-footer)
        --show-date-time <SHOWDATETIME>               (valid for: set-footer)
        --date-time-mode <DATETIMEMODE>               (valid for: set-footer)
        --fixed-date-time-text <FIXEDDATETIMETEXT>    (valid for: set-footer)
        --show-on-title-slide <SHOWONTITLESLIDE>      (valid for: set-footer)
    -o, --output <PATH>                               Write output to file
                                                      instead of stdout. For
                                                      image results, decodes and
                                                      saves as binary file
```

## `pptcli service`

```text
DESCRIPTION:
Start, stop, or check the status of the pptcli background daemon

USAGE:
    pptcli service [OPTIONS] <COMMAND>

OPTIONS:
    -h, --help    Prints help information

COMMANDS:
    start     Start the daemon if it isn't already running
    stop      Stop the running daemon
    status    Report whether the daemon is running
```

### `pptcli service start`

```text
DESCRIPTION:
Start the daemon if it isn't already running

USAGE:
    pptcli service start [OPTIONS]

OPTIONS:
    -h, --help                     Prints help information
        --pipe-name <PIPE_NAME>    Override the daemon's named pipe (defaults to
                                   a per-user pipe name)
```

### `pptcli service status`

```text
DESCRIPTION:
Report whether the daemon is running

USAGE:
    pptcli service status [OPTIONS]

OPTIONS:
    -h, --help                     Prints help information
        --pipe-name <PIPE_NAME>    Override the daemon's named pipe (defaults to
                                   a per-user pipe name)
```

### `pptcli service stop`

```text
DESCRIPTION:
Stop the running daemon

USAGE:
    pptcli service stop [OPTIONS]

OPTIONS:
    -h, --help                     Prints help information
        --pipe-name <PIPE_NAME>    Override the daemon's named pipe (defaults to
                                   a per-user pipe name)
        --force                    Force-kill the daemon process if a graceful
                                   RPC shutdown doesn't respond
```

## `pptcli session`

```text
DESCRIPTION:
Open, create, close, test, Save As/copy, or list presentation sessions held by
the daemon; apply templates, manage the advisory Mark as Final flag, and
read/write document properties

USAGE:
    pptcli session [OPTIONS] <COMMAND>

OPTIONS:
    -h, --help    Prints help information

COMMANDS:
    open <FILE_PATH>                                              Open an
                                                                  existing
                                                                  presentation
                                                                  and return a
                                                                  session id
    create <FILE_PATH>                                            Create a new
                                                                  presentation
                                                                  and return a
                                                                  session id
    close <SESSION_ID>                                            Close a
                                                                  session,
                                                                  optionally
                                                                  saving first
    list                                                          List every
                                                                  session
                                                                  currently open
                                                                  in the daemon
    test <FILE_PATH>                                              Validate that
                                                                  PowerPoint can
                                                                  open a
                                                                  presentation
                                                                  without
                                                                  retaining a
                                                                  session
    save-as <SESSION_ID> <TARGET_PATH>                            Save the
                                                                  active
                                                                  presentation
                                                                  under a new
                                                                  path and move
                                                                  the session to
                                                                  it
    save-copy-as <SESSION_ID> <TARGET_PATH>                       Save a copy
                                                                  without
                                                                  changing the
                                                                  active
                                                                  presentation
                                                                  or session
                                                                  path
    apply-template <SESSION_ID> <TEMPLATE_PATH>                   Apply a
                                                                  template's
                                                                  masters/theme/
                                                                  layouts to the
                                                                  open
                                                                  presentation,
                                                                  preserving
                                                                  slide content
    get-theme-name <SESSION_ID>                                   Read the
                                                                  design/theme
                                                                  name currently
                                                                  applied to the
                                                                  open
                                                                  presentation
    get-final <SESSION_ID>                                        Read
                                                                  PowerPoint's
                                                                  advisory Mark
                                                                  as Final
                                                                  editing flag;
                                                                  it is not
                                                                  authentication
                                                                  , encryption,
                                                                  or access
                                                                  control
    set-final <SESSION_ID> <IS_FINAL>                             Set or clear
                                                                  PowerPoint's
                                                                  advisory Mark
                                                                  as Final
                                                                  editing flag;
                                                                  it is not
                                                                  authentication
                                                                  , encryption,
                                                                  or access
                                                                  control
    set-document-property <SESSION_ID> <PROPERTY_NAME> <VALUE>    Set a built-in
                                                                  document
                                                                  metadata
                                                                  property
                                                                  (Title,
                                                                  Subject,
                                                                  Author,
                                                                  Keywords,
                                                                  Comments,
                                                                  Category,
                                                                  Manager,
                                                                  Company)
    get-document-property <SESSION_ID> <PROPERTY_NAME>            Read a
                                                                  built-in
                                                                  document
                                                                  metadata
                                                                  property
    set-custom-property <SESSION_ID> <PROPERTY_NAME> <VALUE>      Create or
                                                                  update a
                                                                  custom
                                                                  (user-defined)
                                                                  document
                                                                  property
    get-custom-property <SESSION_ID> <PROPERTY_NAME>              Read a custom
                                                                  (user-defined)
                                                                  document
                                                                  property
    remove-custom-property <SESSION_ID> <PROPERTY_NAME>           Remove a
                                                                  custom
                                                                  (user-defined)
                                                                  document
                                                                  property
    set-tag <SESSION_ID> <TAG_NAME> <TAG_VALUE>                   Create or
                                                                  update a
                                                                  case-insensiti
                                                                  ve
                                                                  presentation
                                                                  string tag
    get-tag <SESSION_ID> <TAG_NAME>                               Read a
                                                                  presentation
                                                                  string tag by
                                                                  case-insensiti
                                                                  ve name
    list-tags <SESSION_ID>                                        List
                                                                  presentation
                                                                  string tags in
                                                                  native 1-based
                                                                  order
    delete-tag <SESSION_ID> <TAG_NAME>                            Delete a
                                                                  presentation
                                                                  string tag by
                                                                  case-insensiti
                                                                  ve name
```

### `pptcli session close`

```text
DESCRIPTION:
Close a session, optionally saving first

USAGE:
    pptcli session close <SESSION_ID> [OPTIONS]

ARGUMENTS:
    <SESSION_ID>    Session id returned by 'session open'/'session create'

OPTIONS:
    -h, --help    Prints help information
        --save    Save the presentation before closing it
```

### `pptcli session create`

```text
DESCRIPTION:
Create a new presentation and return a session id

USAGE:
    pptcli session create <FILE_PATH> [OPTIONS]

ARGUMENTS:
    <FILE_PATH>    Full path to the .pptx/.pptm presentation file

OPTIONS:
    -h, --help    Prints help information
```

### `pptcli session get-final`

```text
DESCRIPTION:
Read PowerPoint's advisory Mark as Final editing flag; it is not authentication,
encryption, or access control

USAGE:
    pptcli session get-final <SESSION_ID> [OPTIONS]

ARGUMENTS:
    <SESSION_ID>    Session id returned by 'session open'/'session create'

OPTIONS:
    -h, --help    Prints help information
```

### `pptcli session get-theme-name`

```text
DESCRIPTION:
Read the design/theme name currently applied to the open presentation

USAGE:
    pptcli session get-theme-name <SESSION_ID> [OPTIONS]

ARGUMENTS:
    <SESSION_ID>    Session id returned by 'session open'/'session create'

OPTIONS:
    -h, --help    Prints help information
```

### `pptcli session list`

```text
DESCRIPTION:
List every session currently open in the daemon

USAGE:
    pptcli session list [OPTIONS]

OPTIONS:
    -h, --help    Prints help information
```

### `pptcli session list-tags`

```text
DESCRIPTION:
List presentation string tags in native 1-based order

USAGE:
    pptcli session list-tags <SESSION_ID> [OPTIONS]

ARGUMENTS:
    <SESSION_ID>    Session id returned by 'session open'/'session create'

OPTIONS:
    -h, --help    Prints help information
```

### `pptcli session open`

```text
DESCRIPTION:
Open an existing presentation and return a session id

USAGE:
    pptcli session open <FILE_PATH> [OPTIONS]

ARGUMENTS:
    <FILE_PATH>    Full path to the .pptx/.pptm presentation file

OPTIONS:
    -h, --help    Prints help information
```

### `pptcli session test`

```text
DESCRIPTION:
Validate that PowerPoint can open a presentation without retaining a session

USAGE:
    pptcli session test <FILE_PATH> [OPTIONS]

ARGUMENTS:
    <FILE_PATH>    Full path to the .pptx/.pptm presentation file

OPTIONS:
    -h, --help    Prints help information
```

## `pptcli shape`

```text
DESCRIPTION:
Shape commands: create, inspect, format, group, link, and edit native
placeholders. Operates within an already-open IPresentationBatch, targeting a
specific slide by its 1-based index

USAGE:
    pptcli shape <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                               Prints help information
    -s, --session <SESSION>                  Session ID from 'session open'
                                             command
        --slide-index <SLIDEINDEX>           (required)
        --left <LEFT>                        (required for: add-rectangle,
                                             add-text-box, add-auto-shape,
                                             set-position) (valid for:
                                             add-rectangle, add-text-box,
                                             add-auto-shape, set-position)
        --top <TOP>                          (required for: add-rectangle,
                                             add-text-box, add-auto-shape,
                                             set-position) (valid for:
                                             add-rectangle, add-text-box,
                                             add-auto-shape, set-position)
        --width <WIDTH>                      (required for: add-rectangle,
                                             add-text-box, add-auto-shape,
                                             set-size) (valid for:
                                             add-rectangle, add-text-box,
                                             add-auto-shape, set-size)
        --height <HEIGHT>                    (required for: add-rectangle,
                                             add-text-box, add-auto-shape,
                                             set-size) (valid for:
                                             add-rectangle, add-text-box,
                                             add-auto-shape, set-size)
        --text <TEXT>                        (required for: add-text-box,
                                             set-placeholder-text) (valid for:
                                             add-text-box, set-placeholder-text)
        --shape-type <SHAPETYPE>             (required for: add-auto-shape)
                                             (valid for: add-auto-shape)
        --begin-x <BEGINX>                   (required for: add-line,
                                             add-connector) (valid for:
                                             add-line, add-connector)
        --begin-y <BEGINY>                   (required for: add-line,
                                             add-connector) (valid for:
                                             add-line, add-connector)
        --end-x <ENDX>                       (required for: add-line,
                                             add-connector) (valid for:
                                             add-line, add-connector)
        --end-y <ENDY>                       (required for: add-line,
                                             add-connector) (valid for:
                                             add-line, add-connector)
        --connector-type <CONNECTORTYPE>     (required for: add-connector)
                                             (valid for: add-connector)
        --shape-index <SHAPEINDEX>           (required for: delete,
                                             set-position, set-size, set-fill,
                                             get-fill, set-line, get-line,
                                             set-rotation, get-rotation, flip,
                                             set-z-order, set-shadow,
                                             get-shadow, set-glow, get-glow,
                                             set-reflection, get-reflection,
                                             set-soft-edge, get-soft-edge,
                                             set-bevel, get-bevel, ungroup,
                                             set-name, get-name, set-alt-text,
                                             get-alt-text, set-hyperlink,
                                             get-hyperlink, remove-hyperlink,
                                             get-link-info, update-link,
                                             break-link, set-link-auto-update,
                                             set-placeholder-text,
                                             set-placeholder-image, set-tag,
                                             get-tag, list-tags, delete-tag)
                                             (valid for: delete, set-position,
                                             set-size, set-fill, get-fill,
                                             set-line, get-line, set-rotation,
                                             get-rotation, flip, set-z-order,
                                             set-shadow, get-shadow, set-glow,
                                             get-glow, set-reflection,
                                             get-reflection, set-soft-edge,
                                             get-soft-edge, set-bevel,
                                             get-bevel, ungroup, set-name,
                                             get-name, set-alt-text,
                                             get-alt-text, set-hyperlink,
                                             get-hyperlink, remove-hyperlink,
                                             get-link-info, update-link,
                                             break-link, set-link-auto-update,
                                             set-placeholder-text,
                                             set-placeholder-image, set-tag,
                                             get-tag, list-tags, delete-tag)
        --red <RED>                          (required for: set-fill, set-glow)
                                             (valid for: set-fill, set-line,
                                             set-shadow, set-glow)
        --green <GREEN>                      (required for: set-fill, set-glow)
                                             (valid for: set-fill, set-line,
                                             set-shadow, set-glow)
        --blue <BLUE>                        (required for: set-fill, set-glow)
                                             (valid for: set-fill, set-line,
                                             set-shadow, set-glow)
        --weight <WEIGHT>                    (valid for: set-line)
        --dash-style <DASHSTYLE>             (valid for: set-line)
        --visible <VISIBLE>                  (required for: set-shadow,
                                             set-reflection) (valid for:
                                             set-line, set-shadow,
                                             set-reflection)
        --degrees <DEGREES>                  (required for: set-rotation) (valid
                                             for: set-rotation)
        --direction <DIRECTION>              (required for: flip) (valid for:
                                             flip)
        --z-order-command <ZORDERCOMMAND>    (required for: set-z-order) (valid
                                             for: set-z-order)
        --transparency <TRANSPARENCY>        (valid for: set-shadow, set-glow,
                                             set-reflection)
        --blur <BLUR>                        (valid for: set-shadow,
                                             set-reflection)
        --offset-x <OFFSETX>                 (valid for: set-shadow)
        --offset-y <OFFSETY>                 (valid for: set-shadow)
        --radius <RADIUS>                    (required for: set-glow,
                                             set-soft-edge) (valid for:
                                             set-glow, set-soft-edge)
        --size <SIZE>                        (valid for: set-reflection)
        --bevel-type <BEVELTYPE>             (required for: set-bevel) (valid
                                             for: set-bevel)
        --depth <DEPTH>                      (valid for: set-bevel)
        --inset <INSET>                      (valid for: set-bevel)
        --shape-indexes <SHAPEINDEXES>       (required for: group) (valid for:
                                             group) (JSON format)
        --name <NAME>                        (required for: set-name) (valid
                                             for: set-name)
        --alt-text <ALTTEXT>                 (required for: set-alt-text) (valid
                                             for: set-alt-text)
        --address <ADDRESS>                  (required for: set-hyperlink)
                                             (valid for: set-hyperlink)
        --screen-tip <SCREENTIP>             (valid for: set-hyperlink)
        --auto-update <AUTOUPDATE>           True for automatic refresh or false
                                             for manual refresh. (required for:
                                             set-link-auto-update) (valid for:
                                             set-link-auto-update)
        --image-path <IMAGEPATH>             (required for:
                                             set-placeholder-image) (valid for:
                                             set-placeholder-image)
        --tag-name <TAGNAME>                 (required for: set-tag, get-tag,
                                             delete-tag) (valid for: set-tag,
                                             get-tag, delete-tag)
        --tag-value <TAGVALUE>               (required for: set-tag) (valid for:
                                             set-tag)
    -o, --output <PATH>                      Write output to file instead of
                                             stdout. For image results, decodes
                                             and saves as binary file
```

## `pptcli slide`

```text
DESCRIPTION:
Slide lifecycle, background, section, legacy comment, and slide-import commands

USAGE:
    pptcli slide <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                               Prints help
                                                             information
    -s, --session <SESSION>                                  Session ID from
                                                             'session open'
                                                             command
        --slide-index <SLIDEINDEX>                           (required for:
                                                             delete, duplicate,
                                                             move-to,
                                                             set-background-colo
                                                             r,
                                                             get-background-colo
                                                             r,
                                                             set-gradient-backgr
                                                             ound,
                                                             get-gradient-backgr
                                                             ound,
                                                             list-comments,
                                                             add-comment,
                                                             delete-comment,
                                                             clear-comments,
                                                             set-tag, get-tag,
                                                             list-tags,
                                                             delete-tag) (valid
                                                             for: delete,
                                                             duplicate, move-to,
                                                             set-background-colo
                                                             r,
                                                             get-background-colo
                                                             r,
                                                             set-gradient-backgr
                                                             ound,
                                                             get-gradient-backgr
                                                             ound,
                                                             list-comments,
                                                             add-comment,
                                                             delete-comment,
                                                             clear-comments,
                                                             set-tag, get-tag,
                                                             list-tags,
                                                             delete-tag)
        --to-position <TOPOSITION>                           (required for:
                                                             move-to) (valid
                                                             for: move-to)
        --red <RED>                                          (required for:
                                                             set-background-colo
                                                             r) (valid for:
                                                             set-background-colo
                                                             r)
        --green <GREEN>                                      (required for:
                                                             set-background-colo
                                                             r) (valid for:
                                                             set-background-colo
                                                             r)
        --blue <BLUE>                                        (required for:
                                                             set-background-colo
                                                             r) (valid for:
                                                             set-background-colo
                                                             r)
        --red1 <RED1>                                        (required for:
                                                             set-gradient-backgr
                                                             ound) (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --green1 <GREEN1>                                    (required for:
                                                             set-gradient-backgr
                                                             ound) (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --blue1 <BLUE1>                                      (required for:
                                                             set-gradient-backgr
                                                             ound) (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --red2 <RED2>                                        (required for:
                                                             set-gradient-backgr
                                                             ound) (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --green2 <GREEN2>                                    (required for:
                                                             set-gradient-backgr
                                                             ound) (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --blue2 <BLUE2>                                      (required for:
                                                             set-gradient-backgr
                                                             ound) (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --gradient-style <GRADIENTSTYLE>                     (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --gradient-variant <GRADIENTVARIANT>                 (valid for:
                                                             set-gradient-backgr
                                                             ound)
        --section-index <SECTIONINDEX>                       (required for:
                                                             add-section,
                                                             rename-section,
                                                             delete-section,
                                                             get-section-name)
                                                             (valid for:
                                                             add-section,
                                                             rename-section,
                                                             delete-section,
                                                             get-section-name)
        --section-name <SECTIONNAME>                         (required for:
                                                             rename-section)
                                                             (valid for:
                                                             add-section,
                                                             rename-section)
        --delete-slides <DELETESLIDES>                       (valid for:
                                                             delete-section)
        --author <AUTHOR>                                    (required for:
                                                             add-comment) (valid
                                                             for: add-comment)
        --initials <INITIALS>                                (required for:
                                                             add-comment) (valid
                                                             for: add-comment)
        --text <TEXT>                                        (required for:
                                                             add-comment) (valid
                                                             for: add-comment)
        --left <LEFT>                                        (valid for:
                                                             add-comment)
        --top <TOP>                                          (valid for:
                                                             add-comment)
        --comment-index <COMMENTINDEX>                       (required for:
                                                             delete-comment)
                                                             (valid for:
                                                             delete-comment)
        --source-file-path <SOURCEFILEPATH>                  (required for:
                                                             import-from-file)
                                                             (valid for:
                                                             import-from-file)
        --destination-slide-index <DESTINATIONSLIDEINDEX>    (required for:
                                                             import-from-file)
                                                             (valid for:
                                                             import-from-file)
        --source-start-slide <SOURCESTARTSLIDE>              (valid for:
                                                             import-from-file)
        --source-end-slide <SOURCEENDSLIDE>                  (valid for:
                                                             import-from-file)
        --tag-name <TAGNAME>                                 (required for:
                                                             set-tag, get-tag,
                                                             delete-tag) (valid
                                                             for: set-tag,
                                                             get-tag,
                                                             delete-tag)
        --tag-value <TAGVALUE>                               (required for:
                                                             set-tag) (valid
                                                             for: set-tag)
    -o, --output <PATH>                                      Write output to
                                                             file instead of
                                                             stdout. For image
                                                             results, decodes
                                                             and saves as binary
                                                             file
```

## `pptcli smartart`

```text
DESCRIPTION:
SmartArt commands: add a SmartArt diagram to a slide from PowerPoint's built-in
layout gallery, and add/read/update/delete/count the diagram's nodes. Operates
within an already-open , targeting a specific slide and shape by 1-based index

USAGE:
    pptcli smartart <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                   Prints help information
    -s, --session <SESSION>                      Session ID from 'session open'
                                                 command
        --slide-index <SLIDEINDEX>               (required)
        --layout-name <LAYOUTNAME>               (required for: add-smart-art)
                                                 (valid for: add-smart-art)
        --left <LEFT>                            (required for: add-smart-art)
                                                 (valid for: add-smart-art)
        --top <TOP>                              (required for: add-smart-art)
                                                 (valid for: add-smart-art)
        --width <WIDTH>                          (required for: add-smart-art)
                                                 (valid for: add-smart-art)
        --height <HEIGHT>                        (required for: add-smart-art)
                                                 (valid for: add-smart-art)
        --shape-index <SHAPEINDEX>               (required for: add-node,
                                                 add-child-node, set-node-text,
                                                 get-node-text, delete-node,
                                                 get-node-count) (valid for:
                                                 add-node, add-child-node,
                                                 set-node-text, get-node-text,
                                                 delete-node, get-node-count)
        --text <TEXT>                            (required for: add-node,
                                                 add-child-node, set-node-text)
                                                 (valid for: add-node,
                                                 add-child-node, set-node-text)
        --parent-node-index <PARENTNODEINDEX>    (required for: add-child-node)
                                                 (valid for: add-child-node)
        --node-index <NODEINDEX>                 (required for: set-node-text,
                                                 get-node-text, delete-node)
                                                 (valid for: set-node-text,
                                                 get-node-text, delete-node)
    -o, --output <PATH>                          Write output to file instead of
                                                 stdout. For image results,
                                                 decodes and saves as binary
                                                 file
```

## `pptcli table`

```text
DESCRIPTION:
Table commands: add a table shape, read/write cell text, insert/delete rows and
columns, format cell fill and borders, and merge cells. Operates within an
already-open IPresentationBatch, targeting a specific slide and table shape by
their 1-based indices

USAGE:
    pptcli table <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                               Prints help information
    -s, --session <SESSION>                  Session ID from 'session open'
                                             command
        --slide-index <SLIDEINDEX>           (required)
        --rows <ROWS>                        (required for: add-table) (valid
                                             for: add-table)
        --columns <COLUMNS>                  (required for: add-table) (valid
                                             for: add-table)
        --left <LEFT>                        (required for: add-table) (valid
                                             for: add-table)
        --top <TOP>                          (required for: add-table) (valid
                                             for: add-table)
        --width <WIDTH>                      (required for: add-table) (valid
                                             for: add-table)
        --height <HEIGHT>                    (required for: add-table) (valid
                                             for: add-table)
        --shape-index <SHAPEINDEX>           (required for: set-cell-text,
                                             get-cell-text, insert-row,
                                             delete-row, insert-column,
                                             delete-column, set-cell-fill,
                                             get-cell-fill, set-cell-border,
                                             get-cell-border, merge-cells)
                                             (valid for: set-cell-text,
                                             get-cell-text, insert-row,
                                             delete-row, insert-column,
                                             delete-column, set-cell-fill,
                                             get-cell-fill, set-cell-border,
                                             get-cell-border, merge-cells)
        --row <ROW>                          (required for: set-cell-text,
                                             get-cell-text, delete-row,
                                             set-cell-fill, get-cell-fill,
                                             set-cell-border, get-cell-border,
                                             merge-cells) (valid for:
                                             set-cell-text, get-cell-text,
                                             delete-row, set-cell-fill,
                                             get-cell-fill, set-cell-border,
                                             get-cell-border, merge-cells)
        --column <COLUMN>                    (required for: set-cell-text,
                                             get-cell-text, delete-column,
                                             set-cell-fill, get-cell-fill,
                                             set-cell-border, get-cell-border,
                                             merge-cells) (valid for:
                                             set-cell-text, get-cell-text,
                                             delete-column, set-cell-fill,
                                             get-cell-fill, set-cell-border,
                                             get-cell-border, merge-cells)
        --text <TEXT>                        (required for: set-cell-text)
                                             (valid for: set-cell-text)
        --before-row <BEFOREROW>             (valid for: insert-row)
        --before-column <BEFORECOLUMN>       (valid for: insert-column)
        --red <RED>                          (required for: set-cell-fill)
                                             (valid for: set-cell-fill,
                                             set-cell-border)
        --green <GREEN>                      (required for: set-cell-fill)
                                             (valid for: set-cell-fill,
                                             set-cell-border)
        --blue <BLUE>                        (required for: set-cell-fill)
                                             (valid for: set-cell-fill,
                                             set-cell-border)
        --border-type <BORDERTYPE>           (required for: set-cell-border,
                                             get-cell-border) (valid for:
                                             set-cell-border, get-cell-border)
        --weight <WEIGHT>                    (valid for: set-cell-border)
        --dash-style <DASHSTYLE>             (valid for: set-cell-border)
        --visible <VISIBLE>                  (valid for: set-cell-border)
        --merge-to-row <MERGETOROW>          (required for: merge-cells) (valid
                                             for: merge-cells)
        --merge-to-column <MERGETOCOLUMN>    (required for: merge-cells) (valid
                                             for: merge-cells)
    -o, --output <PATH>                      Write output to file instead of
                                             stdout. For image results, decodes
                                             and saves as binary file
```

## `pptcli textframe`

```text
DESCRIPTION:
Text frame commands: set/get text and basic font formatting (size, bold, italic,
underline, font name, color, alignment, bullets) for a shape's text range.
Operates within an already-open IPresentationBatch, targeting a specific shape
by its 1-based slide and shape index

USAGE:
    pptcli textframe <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                        Prints help information
    -s, --session <SESSION>           Session ID from 'session open' command
        --slide-index <SLIDEINDEX>    (required)
        --shape-index <SHAPEINDEX>    (required)
        --text <TEXT>                 (required for: set-text) (valid for:
                                      set-text)
        --font-size <FONTSIZE>        (required for: set-font-size) (valid for:
                                      set-font-size)
        --bold <BOLD>                 (required for: set-bold) (valid for:
                                      set-bold)
        --red <RED>                   (required for: set-font-color) (valid for:
                                      set-font-color)
        --green <GREEN>               (required for: set-font-color) (valid for:
                                      set-font-color)
        --blue <BLUE>                 (required for: set-font-color) (valid for:
                                      set-font-color)
        --italic <ITALIC>             (required for: set-italic) (valid for:
                                      set-italic)
        --underline <UNDERLINE>       (required for: set-underline) (valid for:
                                      set-underline)
        --font-name <FONTNAME>        (required for: set-font-name) (valid for:
                                      set-font-name)
        --alignment <ALIGNMENT>       (required for: set-alignment) (valid for:
                                      set-alignment)
        --enabled <ENABLED>           (required for: set-bullet) (valid for:
                                      set-bullet)
        --character <CHARACTER>       (valid for: set-bullet)
        --auto-size <AUTOSIZE>        (required for: set-auto-size) (valid for:
                                      set-auto-size)
    -o, --output <PATH>               Write output to file instead of stdout.
                                      For image results, decodes and saves as
                                      binary file
```
