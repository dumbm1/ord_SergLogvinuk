/**
 * ai.jsx (c)MaratShagiev m_js@bk.ru 04.09.2016

 #Adobe ExtendScript epxArtbsToJpgEps.jsx
 ##version: 2.3
 ##compatible: Illustrator CS6+

 ##great destination:
 1. export all artboards to jpg with 300dpi and max quality
 2. resize all artboards
 3. export all artboards to eps

 ##pre conditions:
 1. The Adobe ExtendScript Toolkit application must be installed
 2. All artboards must be the same size, and squard
 3. Objects from different artboards should not overlap with other artboards

 ##using:
 1. open script file, change settings, save the script
 2. open Illustrator file and run the script

 ##peculiarity:
 If the illustrations are very complex and a large number of artboards,
 then the script can take a very long time.
 It may be about 30 minutes, it is depending on the speed of your computer.

 * */
//@target illustrator

/**
 // processed all open documents
 // todo: refactor as reverse loop
 ( function () {
 if ( documents.length == 1 ) {
 expAndScaleEachArtb ( documents[ 0 ] );
 return;
 } else {
 for ( var i = 0; i < documents.length; i++ ) {
 expAndScaleEachArtb ( documents[ i ] );
 i--;
 } } } ());
 */
expArtbsToJpgEps (activeDocument);

function expArtbsToJpgEps (adoc) {

  try {
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    /**
     * BEGIN THE SETTINGS THAT MAY BE CHANGED BY USER
     * */
      // USER CUSTOM FOLDER
    var folderPath         = '/d/кнопки/1';
    // WIDTH IN POINTS FOR EPS SCALING TO
    var epsWidthTo         = 50;
    /**
     * END THE SETTINGS THAT MAY BE CHANGED BY USER
     * */

    var storeInteractLavel = app.userInteractionLevel,
        fileName           = adoc.name.slice (0, adoc.name.lastIndexOf ('.')),
        fullPath           = folderPath + '/' + fileName,
        artbLen            = adoc.artboards.length,
        totalDate          = new Date ();

    (new Folder (folderPath).exists == false) ? new Folder (folderPath).create () : '';

    $.writeln ('>>>==============================>>>');

    runAndLog ([unlockUnhide], ['unlock and unhide'], 1, 0);
    runAndLog ([saveAsPdf], ['save pdf'], 1, 0, fullPath);
    runAndLog ([makeJpgFromPdf], ['pass to photoshop making jpeg from pdf'], 1, 0, fullPath);
    runAndLog ([scaleAndFit], ['scale and fit'], 1, 0);
    runAndLog ([delEmptyLays], ['delete empty layers'], 1, 0);
    runAndLog ([expToEps], ['export to eps'], 1, 0, fullPath);

    $.writeln ('all illustrator processes are completed\ntotal script runtime: ' +
      formatMsec (new Date () - totalDate));
    $.writeln ('<<<==============================<<<');

  } catch (e) {
    alert (e.line + '\n' + e.message);
  } finally {
    //adoc.close ( SaveOptions.DONOTSAVECHANGES );
    app.userInteractionLevel = storeInteractLavel;
  }

  /********************
   *** THE LIBRARY ***
   ******************* */

  function saveAsPdf (fullPath) {
    var pdfSaveOpts = new PDFSaveOptions (),
        f           = new File (fullPath);

    pdfSaveOpts.PDFPreset = '[Illustrator Default]';
    /**
     * COLORCONVERSIONREPURPOSE COLORCONVERSIONTODEST None
     * */
    pdfSaveOpts.colorConversionID = ColorConversion.COLORCONVERSIONREPURPOSE;
    pdfSaveOpts.viewAfterSaving = false;

    adoc.saveAs (f, pdfSaveOpts);
  }

  function makeJpgFromPdf (fullPath) {
//    var btCount = 1;
    sendBt ();

    function sendBt () {
      /* $.writeln (
       '\n' + new Array ( btCount ).join ( '-' ) + '->\n' +
       new Array ( btCount ).join ( '-' ) + '-> ' + fileName + ' sending bt message #' + btCount++ + '\n' +
       new Array ( btCount - 1 ).join ( '-' ) + '->\n'
       );*/
      var bt       = new BridgeTalk ();
      bt.target    = 'photoshop';
      bt.body      = _make.toString () + ';_make("' + fullPath + '","' + artbLen + '");';
      bt.timeout   = 1200;
      bt.onTimeout = function () {
        sendBt ();
      }
      return bt.send ();
    }

    function _make (fullPath, artbLen) {

      app.displayDialogs = DialogModes.NO;

      var pdfFile     = new File (fullPath + '.pdf'),
          jpgPath,
          pdfOpenOpts = new PDFOpenOptions,
          jpgSaveOpts = new JPEGSaveOptions ();

      pdfOpenOpts.usePageNumber        = true;
      pdfOpenOpts.resolution           = 300.0;
      pdfOpenOpts.antiAlias            = true;
      pdfOpenOpts.bitsPerChannel       = BitsPerChannelType.EIGHT;
      pdfOpenOpts.cropPage             = CropToType.CROPBOX;
      pdfOpenOpts.mode                 = OpenDocumentMode.RGB;
      pdfOpenOpts.suppressWarnings     = true;
      pdfOpenOpts.height               = '1000 px';
      pdfOpenOpts.width                = '1000 px';
      pdfOpenOpts.constrainProportions = true;

      jpgSaveOpts.embedColorProfile = false;
      jpgSaveOpts.formatOptions     = FormatOptions.STANDARDBASELINE; // OPTIMIZEDBASELINE PROGRESSIVE STANDARDBASELINE
      jpgSaveOpts.matte             = MatteType.NONE // BACKGROUND BLACK FOREGROUND NETSCAPE NONE SEMIGRAY WHITE
      jpgSaveOpts.quality           = 11; // number [0..12]
//      jpgSaveOpts.scans = 3; // number [3..5] only for when formatOptions = FormatOptions.PROGRESSIVE

      try {
        for (var i = 1; i < artbLen + 1; i++) {
          ( i < 10 ) ? jpgPath = fullPath + '-0' + i : jpgPath = fullPath + '-' + i;
          pdfOpenOpts.page = i;
          app.open (pdfFile, pdfOpenOpts);
          app.activeDocument.saveAs (new File (jpgPath), jpgSaveOpts, true);
          app.activeDocument.close (SaveOptions.DONOTSAVECHANGES);
        }
      } catch (e) {
      } finally {
        pdfFile.remove ();
      }
    }
  }

  function scaleAndFit () {
    var artbWidth = adoc.artboards[0].artboardRect[0] - adoc.artboards[0].artboardRect[2];
    var scaleFact = Math.abs (epsWidthTo * 100 / artbWidth) + '';

    if (!isNum (epsWidthTo) || epsWidthTo < 0 || Math.abs (artbWidth) < epsWidthTo) {
      throw new Error (
        'The value of the variable epsWidthTo must be:\n' +
        ' – number\n' +
        ' – greater than zero\n' +
        ' – less than the width of the artboards\n' +
        'Enter the correct value and try again.'
      );
      return;
    }

    if (!scaleFact.match (/\.\d+$/)) {  // WARN: if no decimal point then add it to string
      scaleFact += '.0';
    }

    for (var j = 0; j < activeDocument.artboards.length; j++) {
      var artb = activeDocument.artboards[j];
      activeDocument.artboards.setActiveArtboardIndex (j);
      executeMenuCommand ('selectallinartboard');
      _act_scale (scaleFact);
      executeMenuCommand ('Fit Artboard to selected Art');
      executeMenuCommand ('deselectall');
    }
    var rws = Math.round (Math.sqrt (adoc.artboards.length));
    activeDocument.rearrangeArtboards (DocumentArtboardLayout.GridByRow, rws, 20, true);

    function _act_scale (scalePctEps) {
      {
        var actStr = '/version 3' +
          '/name [ 8' +
          '	7365745363616c65' +
          ']' +
          '/isOpen 1' +
          '/actionCount 1' +
          '/action-1 {' +
          '	/name [ 8' +
          '		6163745363616c65' +
          '	]' +
          '	/keyIndex 0' +
          '	/colorIndex 0' +
          '	/isOpen 1' +
          '	/eventCount 1' +
          '	/event-1 {' +
          '		/useRulersIn1stQuadrant 0' +
          '		/internalName (adobe_scale)' +
          '		/localizedName [ 5' +
          '			5363616c65' +
          '		]' +
          '		/isOpen 1' +
          '		/isOn 1' +
          '		/hasDialog 1' +
          '		/showDialog 0' +
          '		/parameterCount 5' +
          '		/parameter-1 {' +
          '			/key 1970169453' +
          '			/showInPalette -1' +
          '			/type (boolean)' +
          '			/value 0' +
          '		}' +
          '		/parameter-2 {' +
          '			/key 1818848869' +
          '			/showInPalette -1' +
          '			/type (boolean)' +
          '			/value 1' +
          '		}' +
          '		/parameter-3 {' +
          '			/key 1752136302' +
          '			/showInPalette -1' +
          '			/type (unit real)' +
          '			/value ' + scalePctEps +
          '			/unit 592474723' +
          '		}' +
          '		/parameter-4 {' +
          '			/key 1987339116' +
          '			/showInPalette -1' +
          '			/type (unit real)' +
          '			/value ' + scalePctEps +
          '			/unit 592474723' +
          '		}' +
          '		/parameter-5 {' +
          '			/key 1668247673' +
          '			/showInPalette -1' +
          '			/type (boolean)' +
          '			/value 0' +
          '		}' +
          '	}' +
          '}'
      }

      var f = new File ('~/ScriptAction.aia');
      f.open ('w');
      f.write (actStr);
      f.close ();
      app.loadAction (f);
      f.remove ();
      app.doScript ("actScale", "setScale", false); // action name, set name
      app.unloadAction ("setScale", ""); // set name

    }
  }

  function expToEps (fullPath) {

    var opts         = new EPSSaveOptions (),
        epsFile      = new File (fullPath),
        copy_epsFile = new File ('' + epsFile + '.eps');

    // opts.artboardRange = '1-5';
    opts.cmykPostScript             = false;
    opts.compatibility              = Compatibility.ILLUSTRATOR10;
    opts.compatibleGradientPrinting = false;
    opts.embedAllFonts              = false;
    opts.embedAllFonts              = false;
    /**
     * How transparency should be flattened when saving EPS and
     * Illustrator file formats with compatibility set to
     * versions of Illustrator earlier than Illustrator 10
     * */
    opts.flattenOuput = OutputFlattening.PRESERVEAPPEARANCE; //PRESERVEPATHS
    opts.includeDocumentThumbnails = false;
    opts.overprint                 = PDFOverprint.PRESERVEPDFOVERPRINT; // DISCARDPDFOVERPRINT
    opts.postScript                = EPSPostScriptLevelEnum.LEVEL2;
    opts.preview                   = EPSPreview.None; // BWTIFF COLORTIFF TRANSPARENTCOLORTIFF
    opts.saveMultipleArtboards     = true;

    adoc.saveAs (epsFile, opts);

    if ((opts.saveMultipleArtboards == true && opts.artboardRange == '') && copy_epsFile.exists) {
      copy_epsFile.remove ();
    }

  }

  function delEmptyLays () {

    for (var i = 0; i < activeDocument.layers.length; i++) {
      var lay = activeDocument.layers[i];
      if (_hasSubs (lay)) {
        _delSubs (lay);
      }
      if (_isEmpty (lay) && activeDocument.layers.length > 1) {
        lay.locked == true ? lay.locked = false : '';
        lay.visible == false ? lay.visible = true : '';
        lay.remove ();
        i--;
      }
    }

    /**
     *
     * @param {Object} lay - Layer
     * @return {Object} lay - Layer
     */
    function _delSubs (lay) {
      for (var i = 0; i < lay.layers.length; i++) {
        var thisSubLay = _getSubs (lay)[i];

        if (_isEmpty (thisSubLay)) {
          thisSubLay.locked == true ? thisSubLay.locked = false : '';
          thisSubLay.visible == false ? thisSubLay.visible = true : '';
          thisSubLay.remove ();
          i--;
        }

        if (_hasSubs (thisSubLay)) {
          var parent = _delSubs (thisSubLay);
          if (_isEmpty (parent)) {
            thisSubLay.locked == true ? thisSubLay.locked = false : '';
            thisSubLay.visible == false ? thisSubLay.visible = true : '';
            thisSubLay.remove ();
            i--;
          }
        }
      }
      return lay;
    }

    /**
     *
     * @param  {Object} lay - Layer
     * @return {Boolean}
     */
    function _hasSubs (lay) {
      try {
        return (lay.layers.length > 0);
      } catch (e) {
        return false;
      }
    }

    /**
     *
     * @param  {Object} lay - Layer
     * @return {Boolean}
     */
    function _isEmpty (lay) {
      try {
        return lay.pageItems.length == 0 && lay.layers.length == 0;
      } catch (e) {
        return false;
      }
    }

    /**
     *
     * @param  {Object} lay - Layer
     * @return {Object/Boolean}  layers / false
     */
    function _getSubs (lay) {
      try {
        return lay.layers;
      } catch (e) {
        return false;
      }
    }
  }

  function unlockUnhide () {
    for (var i = 0; i < activeDocument.layers.length; i++) {
      if (activeDocument.layers[i].visible == false) {
        activeDocument.layers[i].visible = true;
      }
      if (activeDocument.layers[i].locked == true) {
        activeDocument.layers[i].locked = false;
      }
    }
    executeMenuCommand ("unlockAll");
    executeMenuCommand ("showAll");
    executeMenuCommand ("deselectall");
  }

  function runAndLog (funcs, descripts, callCounter, startCoount, arg) {
    var i, j,
        percentWeight = 100 / callCounter,
        date          = new Date (), tmpDate;

    for (i = 0; i < funcs.length; i++) {
      $.writeln (descripts[i] + ':');
      if (callCounter > 1) {
        tmpDate = new Date ();
      }

      for (j = startCoount; j < callCounter; j++) {

        if (callCounter == 1) {
          funcs[i] (arg);
        } else {
          funcs[i] (j);
        }
        $.sleep (20);
        $.write (Math.round (percentWeight * j) + '%  ');
      }
      $.writeln ('100%');
      if (tmpDate) {
        $.writeln ('intermediate runtime: ' + formatMsec (new Date () - tmpDate));
      }
    }
    $.writeln ('total function runtime: ' + formatMsec (new Date () - date) + '\n');
  }

  function formatMsec (millisec) {

    var date       = new Date (millisec),
        formatDate =
          ('00' + date.getUTCHours ()).slice (-2) + ':' +
          ('00' + date.getMinutes ()).slice (-2) + ':' +
          ('00' + date.getSeconds ()).slice (-2) + ':' +
          ('000' + date.getMilliseconds ()).slice (-3);
    return formatDate;
  }

  function isNum (n) {
    return !isNaN (parseFloat (n)) && isFinite (n);
  }
}
