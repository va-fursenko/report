<?php
/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 16.03.16
 * Time: 15:34
 */


/** @const Группы айдишников, на которые забиваются строки входных данных */
$GROUP_KEYS = [
    '21 CENTURY',
    'ASET SG',
    'AUTOM. CONT.',
    'BANK OFFICE',
    'BIRTHDAY',
    'BOOK CLUB',
    'BUSINES MAIL',
    'CERT DB MAIL',
    'CONSULTANTS',
    'CONTIN.GRAD.',
    'COURIER',
    'CROSSSELLING',
    'D-T-D',
    'DB MAIL',
    'DOOR-TO-DOOR',
    'EXHIBITION',
    'EXT DB',
    'FOLLOW UP',
    'GOOGLE ADW',
    'INFO-LINE',
    'INTERNET',
    'KIOSK',
    'LCCIEB',
    'LETTERS',
    'MIR KNIGI',
    'ORIFLAME',
    'PHONE',
    'POST OFFICE',
    'PRINT MEDIA',
    'Phone',
    'Print Media',
    'READERS DIG.',
    'REMAIL',
    'RW STATIONS',
    'STUD BY STUD',
    'TEST PHONE',
    'Телеобзвон',
    'REMAIL 1,2,3',
];





$PERMANENT_ROWS = [
    'Телеобзвон',
    'REMAIL 1,2,3'
];




$COL_INDEXES = array_flip(['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']);




Report::$GROUP_KEYS = $GROUP_KEYS;
Report::$PERMANENT_ROWS = $PERMANENT_ROWS;
Report::$COL_INDEXES = $COL_INDEXES;
Report::$CELL_I1 = 26504;
//Report::$ROW_INDEX_DIFF = 4;