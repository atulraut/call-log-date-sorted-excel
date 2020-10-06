package com.reram.callsort;

import android.net.Uri;
import android.os.Bundle;
import android.os.PowerManager;
import android.os.PowerManager.WakeLock;
import android.accounts.Account;
import android.accounts.AccountManager;
import android.app.Activity;
import android.app.DatePickerDialog;
import android.app.Dialog;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.StringTokenizer;

import org.apache.http.protocol.HTTP;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import android.content.Context;
import android.content.Intent;
import android.os.Environment;
import android.telephony.TelephonyManager;
import android.text.format.DateFormat;
import android.util.Log;
import android.view.Gravity;
import android.view.View;
import android.view.View.OnClickListener;
import android.view.ViewGroup;
import android.webkit.MimeTypeMap;
import android.widget.Button;
import android.widget.DatePicker;
import android.widget.ImageView;
import android.widget.TextView;
import android.view.LayoutInflater;
import android.widget.Toast;

// Call log related 
import android.database.Cursor;
import android.graphics.Paint;
import android.provider.CallLog;

public class ReadExcelActivity extends Activity implements OnClickListener{
	private static final String path="/sdcard/reram/callHistory.xls";
	private static final String LOG_TAG = ReadExcelActivity.class.getName();
	public static final String STORAGE_DIR_NAME = "CALSORTXL";
	final Context context = this;
	protected File mFile   = null;
	protected File mDir    = null;
	protected Uri mUri  =   null;
	protected String strStartDate = null;
	protected String strEndDate = null;
	public volatile static String gStrStart = null;
	public volatile static String gStrEnd = null;
	
	Button btnStartDate;
	Button btnEndDate;
	Button btnAbout;
	Button btnWriteExcelButton; 
	Button btnReadExcelButton;
	Button btnSendExcelButton;
	TextView textView0;
	
    TelephonyManager telephonyManager;
    PowerManager powerManager;
    WakeLock wakeLock;
    static final int DATE_START_ID = 0;
    static final int DATE_END_ID = 1;
    
 // variables to save user selected date and time
    public  int sYear,sMonth,sDay;  
    public  int eYear,eMonth,eDay;
    public int gint = 0;
// declare  the variables to Show/Set the date and time when Time and  Date Picker Dialog first appears
    private int mYear, mMonth, mDay; 
    private String acname = null;
    private String actype = null;
    
 // constructor
    public ReadExcelActivity() {
                // Assign current Date and Time Values to Variables
                final Calendar c = Calendar.getInstance();
                mYear = c.get(Calendar.YEAR);
                mMonth = c.get(Calendar.MONTH);
                mDay = c.get(Calendar.DAY_OF_MONTH);
    }
    
    /** Called when the activity is first created. */
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.main);
        
        // get the references of buttons
        btnStartDate=(Button)findViewById(R.id.buttonStartDate); 
        btnEndDate =(Button)findViewById(R.id.buttonEndDate);
        btnAbout  =(Button)findViewById(R.id.aboutId);
        btnWriteExcelButton = (Button)findViewById(R.id.writeExcel);
        btnWriteExcelButton.setOnClickListener(this);
        btnReadExcelButton = (Button)findViewById(R.id.readExcel);
        btnReadExcelButton.setOnClickListener(this);
        btnSendExcelButton = (Button)findViewById(R.id.sendMail);
        btnSendExcelButton.setOnClickListener(this);
        textView0 = (TextView)findViewById(R.id.txtId0);
        
        textView0.setPaintFlags(textView0.getPaintFlags() |   Paint.UNDERLINE_TEXT_FLAG);
        
        // Logic for thin bold line 

        // Get the telephony manager
        telephonyManager = (TelephonyManager) getSystemService(Context.TELEPHONY_SERVICE);
     // Save Power
        powerManager = (PowerManager) getSystemService(POWER_SERVICE);
        wakeLock = powerManager.newWakeLock(PowerManager.PARTIAL_WAKE_LOCK,
                "RERAM-PowerLock");
        
        // Disable all button 
        btnEndDate.setEnabled(false); 
        btnWriteExcelButton.setEnabled(false); 
        btnReadExcelButton.setEnabled(false); 
        btnSendExcelButton.setEnabled(false); 

        // Set ClickListener on btnSelectDate 
        btnStartDate.setOnClickListener(new View.OnClickListener() {
            
            public void onClick(View v) {
                btnEndDate.setEnabled(true); 
                // Show the DatePickerDialog
            	gint = DATE_START_ID;
                showDialog(DATE_START_ID);
            }
        });
        
        // Set ClickListener on btnSelectTime
        btnEndDate.setOnClickListener(new View.OnClickListener() {
            
            public void onClick(View v) {
                // Show the TimePickerDialog
            	 gint = DATE_END_ID;
                 showDialog(DATE_END_ID);
            }
        });

     // Set ClickListener on btnAbout
        btnAbout.setOnClickListener(new View.OnClickListener() {
            
            public void onClick(View v) {
    			// custom dialog
    			final Dialog dialog = new Dialog(context);
    			dialog.setContentView(R.layout.custom);
    			dialog.setTitle("How to Use app !");
     
    			// set the custom dialog components - text, image and button
    			TextView text = (TextView) dialog.findViewById(R.id.text);
    			text.setText(" Start Date : set, e.g. 10 Mar 2014 \n End   Date : set, e.g. 10 Apr 2014 \n Press button Write to Excel Sheet \n Press button e-Mail ExcelSheet	\n \n Contact : reramhi@gmail.com !");
     
	    			Button dialogButton = (Button) dialog.findViewById(R.id.dialogButtonOK);
	    			// if button is clicked, close the custom dialog
	    			dialogButton.setOnClickListener(new OnClickListener() {
	    				@Override
	    				public void onClick(View v) {
	    					dialog.dismiss();
	    				}
	    			});
    			dialog.show();
            }
        });  
    } // onCreate Ends
    
    /* Checks if external storage is available to at least read */
    public boolean isExternalStorageReadable() {
        String state = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(state) ||
            Environment.MEDIA_MOUNTED_READ_ONLY.equals(state)) {
            return true;
        }
        return false;
    }

    @Override
    public void onResume() {
        super.onResume();  // Always call the superclass method first

        // Save Power
        wakeLock.acquire();
    } // onResume Ends
    
    @Override
    protected void onPause() {
        super.onPause();
        wakeLock.release();
    }
    
    @Override
    protected void onDestroy() {
        super.onDestroy();
        deleteExternalStoragePrivateFile(); // Clean up file
    }
    
    // Register  DatePickerDialog listener
    private DatePickerDialog.OnDateSetListener mDateSetListener =
                       new DatePickerDialog.OnDateSetListener() {
                   // the callback received when the user "sets" the Date in the DatePickerDialog
                           public void onDateSet(DatePicker view, int yearSelected,
                                                 int monthOfYear, int dayOfMonth) {
                        	  Log.v (LOG_TAG, "monthOfYear " + monthOfYear);
                        	  if (gint == DATE_START_ID) { 
                        		  sYear = yearSelected;
                        		  sMonth = monthOfYear+1;
                        		  sDay = dayOfMonth;
                        		  if (monthOfYear+1 < 10) {
                               		  gStrStart = "0"+sMonth+"-"+sDay+"-"+sYear;
                        		  } else
                        			  gStrStart = ""+sMonth+"-"+sDay+"-"+sYear;
                              } if (gint == DATE_END_ID) {
                            	  eYear = yearSelected;
                        		  eMonth = monthOfYear+1;
                        		  eDay = dayOfMonth;
                        		  // UI validation : End date is greater than the start date or equal to start date.
                        		  if (eDay < sDay || eMonth < sMonth || eYear < sYear ) {
                            		  LayoutInflater inflater = getLayoutInflater();
                            		  View layout = inflater.inflate(R.layout.toast,
                            		                                 (ViewGroup) findViewById(R.id.toast_layout_root));
                            		  TextView text = (TextView) layout.findViewById(R.id.text);
                            		  text.setText(" End date is equal to or greater than the start date !");
                            		  Toast toast = new Toast(getApplicationContext());
                            		  toast.setGravity(Gravity.CENTER_VERTICAL, 0, 0);
                            		  toast.setDuration(Toast.LENGTH_LONG);
                            		  toast.setView(layout);
                            		  toast.show();
                            		  btnEndDate.setText("Select End Date");
                            		  btnWriteExcelButton.setEnabled(false); 
                            	      btnReadExcelButton.setEnabled(false); 
                            	      btnSendExcelButton.setEnabled(false); 
                        		  } else {
                        			  if (monthOfYear+1 < 10) {
                        				  gStrEnd = "0"+eMonth+"-"+eDay+"-"+eYear;
                        			  } else
                        				  gStrEnd = ""+eMonth+"-"+eDay+"-"+eYear;
                        		        btnWriteExcelButton.setEnabled(true); 
                        		  }
                              }
                              // Set the Selected Date in Select date Button
                              //btnStartDate.setText("Date selected : "+sDay+"-"+sMonth+"-"+sYear);
                              btnStartDate.setText(gStrStart);
                              //btnEndDate.setText("Date selected : "+eDay+"-"+eMonth+"-"+eYear);
                              btnEndDate.setText(gStrEnd);
                           }
                       };
   // Method automatically gets Called when you call showDialog()  method
   @Override
   protected Dialog onCreateDialog(int id) {
       switch (id) {
       case DATE_START_ID:
    	   // create a new DatePickerDialog with values you want to show 
               return new DatePickerDialog(this,
                           mDateSetListener,
                           mYear, mMonth, mDay);
           // create a new TimePickerDialog with values you want to show 
       case DATE_END_ID:
    	// create a new DatePickerDialog with values you want to show 
           return new DatePickerDialog(this,
                       mDateSetListener,
                       mYear, mMonth, mDay);
       }
       return null;
   }
 
    public void onClick(View v) {
        switch (v.getId()) {
        case R.id.writeExcel:
            writeExcelFile(this,"callHistory.xls");
            break;
        case R.id.readExcel:
            readExcelFile(this,"callHistory.xls");
            break;
        case R.id.sendMail:
            sendMail(this, "callHistory.xls");           
        	break;
        }
    }
 
    private boolean writeExcelFile(Context context, String fileName) { 
 
        // check if available and not read only 
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) { 
            Log.w("FileUtils", "Storage not available or read only"); 
            //return false; 
        } 
 
        boolean success = false; 
 
        //New Workbook
        Workbook wb = (Workbook) new HSSFWorkbook();
 
        Cell c = null;
 
        //Cell style for header row
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.LIME.index);
        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        
        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("callLog");
 
        // Generate column headings
        Row row  = sheet1.createRow(0);
        Row row1 = sheet1.createRow(1);
 
        c = row.createCell(0);
        c.setCellValue("Date");
        c.setCellStyle(cs);
 
        c = row.createCell(1);
        c.setCellValue("Time");
        c.setCellStyle(cs);
 
        c = row.createCell(2);
        c.setCellValue("Number");
        c.setCellStyle(cs);
 
        c = row.createCell(3);
        c.setCellValue("Name"); // Charges(Rs)
        c.setCellStyle(cs);
        
        c = row.createCell(4);
        c.setCellValue("Type");
        c.setCellStyle(cs);

        c = row.createCell(5);
        c.setCellValue("Call duration(min:sec)");
        c.setCellStyle(cs);

        c = row.createCell(6);
        c.setCellValue("Call duration(Seconds)");
        c.setCellStyle(cs);

        c = row.createCell(7);
        c.setCellValue("Call duration(Minutes)");
        c.setCellStyle(cs);

        sheet1.setColumnWidth(0, (15 * 250));
        sheet1.setColumnWidth(1, (15 * 100));
        sheet1.setColumnWidth(2, (15 * 250));
        sheet1.setColumnWidth(3, (15 * 300));
        sheet1.setColumnWidth(4, (15 * 350));
        sheet1.setColumnWidth(5, (15 * 350));
        sheet1.setColumnWidth(6, (15 * 350));
        sheet1.setColumnWidth(7, (15 * 350));
        

       	updateRowContent(sheet1, c);

       	// SD Card File handling.
        // Create a path where we will place our List of objects on external storage 
       	//File reramDirectory = new File("/sdcard/reram/");
	    // have the object build the directory structure, if needed.
       	//reramDirectory.mkdir();
	    // create a File object for the output file
	    //File file = new File(reramDirectory, fileName);
	    File file = null;
        FileOutputStream os = null; 
        boolean chkStorage = false;
        
        try { 
        	chkStorage = isExternalStorageAvailable ();
            Log.w("[createExternalStoragePrivateFile]", "ATUL 22" + chkStorage); 
        	if (chkStorage == true) {
        		file =  createExternalStoragePrivateFile(fileName);
        		os = new FileOutputStream(file);
                //Log.w("[createExternalStoragePrivateFile]", "ATUL 00" + file); 
        	} else {
        		file = createInternalStoragePrivateFile (fileName);
        		os = openFileOutput(fileName, Context.MODE_WORLD_READABLE);
                //Log.w("[createInternalStoragePrivateFile]", "ATUL 11" + file); 
        	}
            wb.write(os);
            //Log.w("FileUtils", "Writing file" + file); 
            success = true; 
        } catch (IOException e) { 
            Log.w("FileUtils", "Error writing " + file, e); 
        } catch (Exception e) { 
            Log.w("FileUtils", "Failed to save file", e); 
        } finally { 
            try { 
                if (null != os) 
                    os.close(); 
            } catch (Exception ex) { 
            } 
        } 
        
        btnEndDate.setEnabled(false); 
        btnWriteExcelButton.setEnabled(false); 
        btnReadExcelButton.setEnabled(true); 
        btnSendExcelButton.setEnabled(true); 
        return success; 
    } 
 
    private void readExcelFile(Context context, String filename) { 
    	/*
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) 
        { 
        	LayoutInflater inflater = getLayoutInflater();
  		  	View layout = inflater.inflate(R.layout.toast,
  		                                 (ViewGroup) findViewById(R.id.toast_layout_root));
  		  	TextView text = (TextView) layout.findViewById(R.id.text);
  		  	text.setText(" Storage not available or read only !");
  		  	Toast toast = new Toast(getApplicationContext());
  		  	toast.setGravity(Gravity.CENTER_VERTICAL, 0, 0);
  		  	toast.setDuration(Toast.LENGTH_LONG);
  		  	toast.setView(layout);
  		  	toast.show();
  		  	Log.w("FileUtils", "Storage not available or read only"); 
            return; 
        } */
 
        try{        	
        	// Creating Input Stream
        	File file2 = null;
        	boolean chkStorage = isExternalStorageAvailable ();
        	if (chkStorage == true) {
        		file2 =  createExternalStoragePrivateFile(filename);
        	} else {
        		file2 = createInternalStoragePrivateFile (filename);
        	}
        	Log.w("[ReadFile]", "ATUL 0000 --> " + file2); 
            //File file = new File(context.getExternalFilesDir("/sdcard/reram/"), filename); 
            //FileInputStream myInput = new FileInputStream(file);
            //String path="/sdcard/reram/callHistory.xls";

            Intent intent = new Intent();
            intent.setAction(android.content.Intent.ACTION_VIEW);
            //File file2 = new File(path);
           
            MimeTypeMap mime = MimeTypeMap.getSingleton();
            String ext=file2.getName().substring(file2.getName().indexOf(".")+1);
            String type = mime.getMimeTypeFromExtension(ext);
          
            intent.setDataAndType(Uri.fromFile(file2),type);
   
            startActivity(intent);
            
            /*
            // Create a POIFSFileSystem object 
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
 
            // Create a workbook using the File System 
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
 
            // Get the first sheet from workbook 
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);
 
            // We now need something to iterate through the cells.
            Iterator<Row> rowIter = mySheet.rowIterator();
            // Read Logic
            while(rowIter.hasNext()){
                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator<Cell> cellIter = myRow.cellIterator();
                while(cellIter.hasNext()){
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    Log.w("FileUtils", "Cell Value: " +  myCell.toString());
                    Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
                }
            } */            
            // Delete unwanted cell/row            
        }catch (Exception e){e.printStackTrace(); }
        return;
    } 
 
    public static boolean isExternalStorageReadOnly() { 
        String extStorageState = Environment.getExternalStorageState(); 
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) { 
            return true; 
        } 
        return false; 
    } 
 
    public static boolean isExternalStorageAvailable() { 
        String extStorageState = Environment.getExternalStorageState(); 
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) { 
            return true; 
        } 
        return false; 
    }
    
    // if external Storage available e.g. SD Card
    public File createExternalStoragePrivateFile(String fileName) {
        // Get the directory for the user's public pictures directory. 
//        File afile = new File(Environment.getExternalStoragePublicDirectory(
//                Environment.DIRECTORY_DOWNLOADS), "RERAM");
        File afile = new File(Environment.getExternalStorageDirectory(), STORAGE_DIR_NAME);        
        File reramDirectory = new File(afile+"");
	    // have the object build the directory structure, if needed.
       	reramDirectory.mkdir();
	    // create a File object for the output file
	    File file = new File(reramDirectory, fileName);
        return file;
    }
    
    // if not external storage then go with internal storage
    public File createInternalStoragePrivateFile (String fileName) {
    	File file = new File(context.getFilesDir(), fileName);
    	return file;
    }
    
    // Fix it not working.
    void deleteExternalStoragePrivateFile() {
        // Get path for the file on external storage.  If external
        // storage is not currently mounted this will fail.
        File afile = new File(Environment.getExternalStorageDirectory(), STORAGE_DIR_NAME);        
        File file = new File(afile+"");
        if (file != null) {
            file.delete();
        }
    }
    
    public static Long createDate(int year, int month, int day)
    {
        Calendar calendar = Calendar.getInstance();

        calendar.set(year, month, day);
        Log.v (LOG_TAG, "ATUL" + calendar.getTimeInMillis());
        return calendar.getTimeInMillis();

    }
    
    protected void updateRowContent(Sheet mySheet2, Cell c2) { 
       	String[] wordStart = gStrStart.split("-"); 
       	/*for (String word : wordStart) {  		  
    		   Log.v (TAG, "ATUL word  = " + word);
    		   Log.v (TAG, "ATUL YEAR word[2]   = " + word[2]);
    		   Log.v (TAG, "ATUL MONTH word[0]  = " + word[0]);
    		   Log.v (TAG, "ATUL DAY word[1]    = " + word[1]);
    	}*/
		int sYear = Integer.parseInt(wordStart[2]);
		int sMonth = Integer.parseInt(wordStart[0]);
		sMonth = sMonth - 1;
		int sDay = Integer.parseInt(wordStart[1]);
		Log.v (LOG_TAG, "ATUL wordStart  = " + sYear + sMonth + sDay);

		Calendar cal = Calendar.getInstance();
//    		cal.set(2014, Calendar.MARCH, 12, 12, 0, 0); // start date: public final void set (int year, int month, int day, int hourOfDay, int minute, int second)
		cal.set(sYear, sMonth, sDay, 0, 0, 0); // start date
		Long fterDate = cal.getTimeInMillis(); 
		String afterDate = fterDate.toString();

		String[] wordEnd = gStrEnd.split("-"); 
		/*for (String word : wordEnd) {  		  
		   Log.v (TAG, "ATUL word  = " + word);
		}*/
		int eYear = Integer.parseInt(wordEnd[2]);
		int eMonth = Integer.parseInt(wordEnd[0]);
		eMonth = eMonth - 1;
		int eDay = Integer.parseInt(wordEnd[1]);
		Log.v (LOG_TAG, "ATUL wordEnd  = " + eYear + eMonth + eDay);
	
//    		cal.set(2014, Calendar.APRIL, 01, 12, 0, 0); // end date:  public final void set (int year, int month, int day, int hourOfDay, int minute, int second)
		cal.set(eYear, eMonth, eDay, 24, 0, 0); // end date
		Long foreDate = cal.getTimeInMillis(); 
		String beforeDate = foreDate.toString();

		Log.v (LOG_TAG, "ATUL afterDate1  = " + afterDate);
	    Log.v (LOG_TAG, "ATUL beforeDate1 = " + beforeDate);
		
/*	    cal.set(2014, Calendar.MARCH, 24, 12, 0, 0); // start date
		fterDate = cal.getTimeInMillis(); 
		afterDate = fterDate.toString();
		cal.set(2014, Calendar.APRIL, 01, 12, 0, 0); // end date
		foreDate = cal.getTimeInMillis(); 
		beforeDate = foreDate.toString();
*/
		Log.v (LOG_TAG, "ATUL 11 afterDate1  = " + afterDate);
	    Log.v (LOG_TAG, "ATUL 11 beforeDate1 = " + beforeDate);

	final String SELECT = CallLog.Calls.DATE + ">?" + " AND "
		+ CallLog.Calls.DATE + "<?";

	Cursor managedCursor = managedQuery(
		CallLog.Calls.CONTENT_URI,
		null,
		SELECT,
		new String[] { afterDate,
		               beforeDate },
		CallLog.Calls.DATE + " desc");

    	int date = managedCursor.getColumnIndex(CallLog.Calls.DATE); 
    	int type = managedCursor.getColumnIndex(CallLog.Calls.TYPE); // The type of the call (incoming[1], outgoing[2] or missed[3]).
    	int number = managedCursor.getColumnIndex(CallLog.Calls.NUMBER); 
    	int duration = managedCursor.getColumnIndex(CallLog.Calls.DURATION); 
    	int cachedname = managedCursor.getColumnIndex(CallLog.Calls.CACHED_NAME);
    	
    	String phNumber = null;
    	String callType = null;
    	String mcallDate = null;
    	Date callDayTime = null;
    	String callDuration = null;
    	String strCachedName = null;
    	String strOperatorName = null;
 
    	SimpleDateFormat datePattern = null;
    	Long datelong = null;
    	String callDate = null;
    	String operatorName = null;
		Calendar calNow = Calendar.getInstance();
		datePattern = new SimpleDateFormat ("MM-dd-yyyy");
    	
    	while (managedCursor.moveToNext()) { 
    		phNumber = managedCursor.getString(number); // Phone Number
    		callType = managedCursor.getString(type); 
    		mcallDate = managedCursor.getString(date); // Date  /*Logic to format in to DateFormate*/
    		datelong = Long.parseLong(mcallDate);
    		callDate = datePattern.format(new Date(datelong));
    		
    		Log.v (LOG_TAG, "ATUL date = " + callDate);
    		callDayTime = new Date(Long.valueOf(mcallDate)); // Time
    		/* Calculate Time when Call came */
    		//Calendar calNow = Calendar.getInstance();
    		calNow.setTime(callDayTime);
    		int hrs = calNow.get(calNow.HOUR_OF_DAY);    		
    		//Log.v (TAG, "ATUL HRS = " + hrs);
    		/* Working Part --- End --- */
    		
    		callDuration = managedCursor.getString(duration); // Call duration
       		
    		strCachedName = managedCursor.getString(cachedname); // Caller Name
    		
    		/*Logic to get time in hr:mm:sec format*/
    		 float caldu =Float.parseFloat(callDuration); //convert seconds into minutes eg. 4secs to 1 minute
             float value = caldu/60;
             float mod = caldu%60;
             String tempStr=""+value;
             String tempStrmod = ""+mod;
             
             StringTokenizer tokens = new StringTokenizer(tempStr, ".");
             String strToken1=tokens.nextToken();
             String strToken2=tokens.nextToken();
             int lVal=Integer.parseInt(strToken1);
             int rVal=Integer.parseInt(strToken2);
             String CallsDurationStrMin = null;
             String MinSec = ":";
             if(rVal>0) {
                 lVal=lVal+1;
                 CallsDurationStrMin=""+lVal;
             } else if(rVal==0) {
                 CallsDurationStrMin=""+lVal;
             }
             tokens = new StringTokenizer(tempStrmod, ".");
             strToken2=tokens.nextToken();
             MinSec = MinSec.concat(strToken2);
             MinSec = strToken1.concat(MinSec);
             
             // [callType] to get Operator Name
             int dircode = Integer.parseInt(callType);
             switch (dircode) {
             case CallLog.Calls.OUTGOING_TYPE:
            	 strOperatorName = "OUTGOING";
                 break;

             case CallLog.Calls.INCOMING_TYPE:
            	 strOperatorName = "INCOMING";
                 break;

             case CallLog.Calls.MISSED_TYPE:
            	 strOperatorName = "MISSED";
                 break;
             }

             //Log.w("FileUtils", "[updateRowContent] 1 :" + CallsDurationStrMin);             
             //Log.w("FileUtils", "[updateRowContent] 2 :" + MinSec); 
             //sb.append("\nPhone Number:--- " + phNumber + " \nCall Type:--- " + callType + " \nCall Date:--- " + callDayTime + " \nCall duration in sec :--- " + callDuration); 
             //sb.append("\n----------------------------------");

    		int lastIndex=mySheet2.getLastRowNum();
    		int i=lastIndex+1;
    		Row row = mySheet2.createRow(i);
    		// Update excel according to set Date
    		//if (callDate.equals(gStrStart) || callDate.equals(gStrEnd)) {
	    		c2 = row.createCell(0);
	    		c2.setCellValue(callDate); // Date
	    		
	    		c2 = row.createCell(1);
	    		c2.setCellValue(hrs); // Time [callDayTime]
	
	    		c2 = row.createCell(2);
	    		c2.setCellValue(phNumber); // Number

	    		c2 = row.createCell(3);
	    		c2.setCellValue(strCachedName); // Name
	
	    		c2 = row.createCell(4);
	    		c2.setCellValue(strOperatorName); // Operator [callType]
	    		
	    		c2 = row.createCell(5);
	    		c2.setCellValue(MinSec); // Call Duration (min:sec)
	    		
	    		c2 = row.createCell(6);
	    		c2.setCellValue(callDuration); // Call Duration (Seconds)
	
	    		c2 = row.createCell(7);
	    		c2.setCellValue(CallsDurationStrMin); // Call Duration (Minute)
	
    		//}
    		// Call duration
    	} // while loop end
    }
    
    protected String getGMailAcntId () {
    	AccountManager am = AccountManager.get(context);
        Account[] accounts = am.getAccounts();
        ArrayList<String> googleAccounts = new ArrayList<String>();
        for (Account ac : accounts) {
            acname = ac.name;
            actype = ac.type;
            //add only google accounts
            if(ac.type.equals("com.google")) {
                googleAccounts.add(ac.name);
                Log.d(LOG_TAG, "RAJVEER: " + acname);
            }
            Log.d(LOG_TAG, "accountInfo: " + acname + ":" + actype);
        }
        return acname;
    }
    protected void sendMail (Context context, String fileName) {	
    	if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) 
        { 
            Log.w("FileUtils", "[sendMail]Storage not available or read only"); 
            return; 
        }
    	// Creating Input Stream 
        //File file = new File(context.getExternalFilesDir("/sdcard/reram/"), fileName); 
        //String path="/sdcard/reram/callHistory.xls";
        File file = new File(path); 
         
		if (null != file)
			mUri  =   Uri.fromFile(file);
	    /*  
		Intent sendIntent = new Intent(Intent.ACTION_SEND);
		sendIntent.putExtra(Intent.EXTRA_SUBJECT, "Person Details");
		sendIntent.putExtra(Intent.EXTRA_STREAM, u1);
		sendIntent.setType("text/html");
		startActivity(sendIntent);
	    */
		String strGAcntId = getGMailAcntId ();
	   /* Send e-mail */
		Intent emailIntent = new Intent(Intent.ACTION_SEND);
		// The intent does not have a URI, so declare the "text/plain" MIME type
		emailIntent.setType(HTTP.PLAIN_TEXT_TYPE);
		emailIntent.putExtra(Intent.EXTRA_EMAIL, new String[] {strGAcntId}); // recipients
		emailIntent.putExtra(Intent.EXTRA_SUBJECT, "Android: Call Log/History");
		emailIntent.putExtra(Intent.EXTRA_STREAM, mUri);
		emailIntent.putExtra(Intent.EXTRA_TEXT, "Email message Excel attach");
		emailIntent.putExtra(Intent.EXTRA_STREAM, mUri);
		this.startActivity(emailIntent);
//        this.startActivity(Intent.createChooser(emailIntent, "E-mail"));
    } // end sendMail function
}
