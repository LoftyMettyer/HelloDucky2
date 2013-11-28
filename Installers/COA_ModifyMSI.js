// ModifyMSI.js 
// Performs a post-build fixup of an msi to change certain properties

// Constant values from Windows Installer
var msiOpenDatabaseModeTransact = 1;

var msiViewModifyInsert = 1
var msiViewModifyUpdate = 2
var msiViewModifyAssign = 3
var msiViewModifyReplace = 4
var msiViewModifyDelete = 6

var msidbCustomActionTypeInScript = 0x00000400;
var msidbCustomActionTypeNoImpersonate = 0x00000800

var msiFontName = 'Verdana';
var msiFontSizeNormal = 8;
var msiFontSizeLarge = 10;
var msiFontColor = 16777215;





if (WScript.Arguments.Length != 1) {
  WScript.StdErr.WriteLine(WScript.ScriptName + " file");
  WScript.Quit(1);
}

var filespec = WScript.Arguments(0);
var installer = WScript.CreateObject("WindowsInstaller.Installer");
var database = installer.OpenDatabase(filespec, msiOpenDatabaseModeTransact);

var sql
var view
var record



try {
  // Delete the CustomAction 'WEBCA_SetTARGETAPPPOOL' to remove the Application Pool combo
  sql = "DELETE FROM `CustomAction` WHERE `CustomAction`.`Action` = 'WEBCA_SetTARGETAPPPOOL'";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();
}
catch (e) {
}



try {
  sql = "SELECT `TextStyle`, `FaceName`, `Size`, `Color`, `StyleBits` FROM `TextStyle`";
  view = database.OpenView(sql);
  view.Execute();
  record = view.Fetch();
  while (record) {
    record.StringData(2) = msiFontName;
    record.IntegerData(5) = 0;
    if (record.StringData(1) == 'VSI_MS_Sans_Serif16.0_1_0') {
      record.IntegerData(3) = msiFontSizeLarge;
//      record.IntegerData(4) = msiFontColor;
    } else {
      record.IntegerData(3) = msiFontSizeNormal;
      record.IntegerData(4) = 0;
    }
    view.Modify(msiViewModifyReplace, record);
    record = view.Fetch();
  }
  view.Close();
  database.Commit();
}
catch (e) {
}



try {
  sql = "SELECT `Control`, `Text` FROM `Control`";
  view = database.OpenView(sql);
  view.Execute();
  record = view.Fetch();
  while (record) {
    if (record.StringData(1) == 'BannerText') {
      record.StringData(2) = record.StringData(2).replace('Welcome to the ', '');
      record.StringData(2) = record.StringData(2).replace(' Wizard', '');
      view.Modify(msiViewModifyReplace, record);
    }
    if (record.StringData(1) == 'Heading') {
      record.StringData(2) = record.StringData(2).replace('Welcome to the ', '');
      record.StringData(2) = record.StringData(2).replace(' Wizard', '');
      view.Modify(msiViewModifyReplace, record);
    }
    if (record.StringData(1) == 'Body1') {
      record.StringData(2) = record.StringData(2).replace('The installer will install [ProductName] to the following web location.', '[ProductName] will be installed to the following web location.');
      record.StringData(2) = record.StringData(2).replace('To install in this folder, click "Next". To install to a different folder, enter it below or click "Browse".','Click "Next" to install to this folder or "Browse" to select to a different folder.');
      view.Modify(msiViewModifyReplace, record);
    }
    record = view.Fetch();
  }
  view.Close();
  database.Commit();
}
catch (e) {
}



try {
  sql = "INSERT INTO `Property` (`Property`.`Property`,`Property`.`Value`) VALUES ('ALLUSERS',2)";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();
}
catch (e) {
}



try {
  sql = "INSERT INTO `Property` (`Property`.`Property`,`Property`.`Value`) VALUES ('ARPCONTACT','Advanced Business Solutions')";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();

  sql = "INSERT INTO `Property` (`Property`.`Property`,`Property`.`Value`) VALUES ('ARPHELPTELEPHONE','08451 609 999')";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();

  sql = "INSERT INTO `Property` (`Property`.`Property`,`Property`.`Value`) VALUES ('ARPHELPLINK','http://webfirst.advancedcomputersoftware.com')";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();

  sql = "INSERT INTO `Property` (`Property`.`Property`,`Property`.`Value`) VALUES ('ARPURLINFOABOUT','http://www.advancedcomputersoftware.com/abs')";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();
}
catch (e) {
}



try {
  sql = "UPDATE `Shortcut` SET `Shortcut`.`WkDir` = 'TARGETDIR'";
  view = database.OpenView(sql);
  view.Execute();
  view.Close();
  database.Commit();
}
catch (e) {
}



try {
  sql = "SELECT `Action`, `Type`, `Source`, `Target` FROM `CustomAction`";
  view = database.OpenView(sql);
  view.Execute();
  record = view.Fetch();
  while (record) {
    if (record.IntegerData(2) & msidbCustomActionTypeInScript) {
      record.IntegerData(2) = record.IntegerData(2) | msidbCustomActionTypeNoImpersonate;
      view.Modify(msiViewModifyReplace, record);
    }
    record = view.Fetch();
  }
  view.Close();
  database.Commit();
}
catch(e) {
}
