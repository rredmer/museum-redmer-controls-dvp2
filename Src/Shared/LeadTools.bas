Attribute VB_Name = "LeadTools"

Global Const ERROR_FAILURE = 20000
Global Const ERROR_NO_MEMORY = 20001
Global Const ERROR_NO_BITMAP = 20002
Global Const ERROR_MEMORY_TOO_LOW = 20003
Global Const ERROR_FILE_LSEEK = 20004
Global Const ERROR_FILE_WRITE = 20005
Global Const ERROR_FILE_GONE = 20006
Global Const ERROR_FILE_READ = 20007
Global Const ERROR_INV_FILENAME = 20008
Global Const ERROR_FILE_FORMAT = 20009
Global Const ERROR_FILENOTFOUND = 20010
Global Const ERROR_INV_RANGE = 20011
Global Const ERROR_IMAGE_TYPE = 20012
Global Const ERROR_INV_PARAMETER = 20013
Global Const ERROR_FILE_OPEN = 20014
Global Const ERROR_UNKNOWN_COMP = 20015
Global Const ERROR_FEATURE_NOT_SUPPORTED = 20016
Global Const ERROR_NOT_256_COLOR = 20017
Global Const ERROR_PRINTER = 20018
Global Const ERROR_CRC_CHECK = 20019
Global Const ERROR_QFACTOR = 20021
Global Const ERROR_TARGAINSTALL = 20022
Global Const ERROR_OUTPUTTYPE = 20023
Global Const ERROR_XORIGIN = 20024
Global Const ERROR_YORIGIN = 20025
Global Const ERROR_VIDEOTYPE = 20026
Global Const ERROR_BITPERPIXEL = 20027
Global Const ERROR_WINDOWSIZE = 20028
Global Const ERROR_NORMAL_ABORT = 20029
Global Const ERROR_NOT_INITIALIZED = 20030
Global Const ERROR_CU_BUSY = 20031
Global Const ERROR_INVALID_TABLE_TYPE = 20032
Global Const ERROR_UNEQUAL_TABLES = 20033
Global Const ERROR_INVALID_BUFFER = 20034
Global Const ERROR_MISSING_TILE_DATA = 20035
Global Const ERROR_INVALID_QVALUE = 20036
Global Const ERROR_INVALIDDATA = 20037
Global Const ERROR_INVALID_COMPRESSED_TYPE = 20038
Global Const ERROR_INVALID_COMPONENT_NUM = 20039
Global Const ERROR_INVALID_PIXEL_TYPE = 20040
Global Const ERROR_INVALID_PIXEL_SAMPLING = 20041
Global Const ERROR_INVALID_SOURCE_FILE = 20042
Global Const ERROR_INVALID_TARGET_FILE = 20043
Global Const ERROR_INVALID_IMAGE_DIMS = 20044
Global Const ERROR_INVALID_TILE_DIMS = 20045
Global Const ERROR_INVALID_PIX_BUFF_DIMS = 20046
Global Const ERROR_SEGMENT_OVERFLOW = 20047
Global Const ERROR_INVALID_SUBSAMPLING = 20048
Global Const ERROR_INVALID_Q_VIS_TABLE = 20049
Global Const ERROR_INVALID_DC_CODE_TABLE = 20050
Global Const ERROR_INVALID_AC_CODE_TABLE = 20051
Global Const ERROR_INSUFFICIENT_DATA = 20052
Global Const ERROR_MISSING_FUNC_POINTER = 20053
Global Const ERROR_TOO_MANY_DC_CODE_TABLES = 20054
Global Const ERROR_TOO_MANY_AC_CODE_TABLES = 20055
Global Const ERROR_INVALID_SUBIMAGE = 20056
Global Const ERROR_INVALID_ABORTION = 20057
Global Const ERROR_CU_NO_SUPPORT = 20058
Global Const ERROR_CU_FAILURE = 20059
Global Const ERROR_BAD_POINTER = 20060
Global Const ERROR_HEADER_DATA_FAILURE = 20061
Global Const ERROR_COMPRESSED_DATA_FAILURE = 20062

Global Const ERROR_FIXEDPAL_DATA = 20063
Global Const ERROR_LOADFONT_DATA = 20064
Global Const ERROR_NO_STAMP = 20065

Global Const ERROR_G3CODE_INVALID = 20070
Global Const ERROR_G3CODE_EOF = 20071
Global Const ERROR_G3CODE_EOL = 20072
Global Const ERROR_PREMATURE_EOF = 20073
Global Const ERROR_PREMATURE_EOL = 20074
Global Const ERROR_UNCOMP_EOF = 20075
Global Const ERROR_ACT_INCOMP = 20076
Global Const ERROR_BAD_DECODE_STATE = 20077
Global Const ERROR_VERSION_NUMBER = 20078
Global Const ERROR_TWAIN_NODSM = 20079
Global Const ERROR_TWAIN_BUMMER = 20080
Global Const ERROR_TWAIN_LOWMEMORY = 20081
Global Const ERROR_TWAIN_NODS = 20082
Global Const ERROR_TWAIN_MAXCONNECTIONS = 20083
Global Const ERROR_TWAIN_OPERATIONERROR = 20084
Global Const ERROR_TWAIN_BADCAP = 20085
Global Const ERROR_TWAIN_BADPROTOCOL = 20086
Global Const ERROR_TWAIN_BADVALUE = 20087
Global Const ERROR_TWAIN_SEQERROR = 20088
Global Const ERROR_TWAIN_BADDEST = 20089
Global Const ERROR_TWAIN_CANCEL = 20090
Global Const ERROR_PANWINDOW_NOT_CREATED = 20091
Global Const ERROR_NOT_ENOUGH_IMAGES = 20092
Global Const ERROR_USER_ABORT = 20100
Global Const ERROR_FPX_INVALID_FORMAT_ERROR = 20101
Global Const ERROR_FPX_FILE_WRITE_ERROR = 20102
Global Const ERROR_FPX_FILE_READ_ERROR = 20103
Global Const ERROR_FPX_FILE_NOT_FOUND = 20104
Global Const ERROR_FPX_COLOR_CONVERSION_ERROR = 20105
Global Const ERROR_FPX_SEVER_INIT_ERROR = 20106
Global Const ERROR_FPX_LOW_MEMORY_ERROR = 20107
Global Const ERROR_FPX_IMAGE_TOO_BIG_ERROR = 20108
Global Const ERROR_FPX_INVALID_COMPRESSION_ERROR = 20109
Global Const ERROR_FPX_INVALID_RESOLUTION = 20110
Global Const ERROR_FPX_INVALID_FPX_HANDLE = 20111
Global Const ERROR_FPX_TOO_MANY_LINES = 20112
Global Const ERROR_FPX_BAD_COORDINATES = 20113
Global Const ERROR_FPX_FILE_SYSTEM_FULL = 20114
Global Const ERROR_FPX_MISSING_TABLE = 20115
Global Const ERROR_FPX_RETURN_PARAMETER_TOO_LARGE = 20116
Global Const ERROR_FPX_NOT_A_VIEW = 20117
Global Const ERROR_FPX_VIEW_IS_TRANFORMLESS = 20118
Global Const ERROR_FPX_ERROR = 20119
Global Const ERROR_FPX_UNIMPLEMENTED_FUNCTION = 20120
Global Const ERROR_FPX_INVALID_IMAGE_DESC = 20121
Global Const ERROR_FPX_INVALID_JPEG_TABLE = 20122
Global Const ERROR_FPX_ILLEGAL_JPEG_ID = 20123
Global Const ERROR_FPX_MEMORY_ALLOCATION_FAILED = 20124
Global Const ERROR_FPX_NO_MEMORY_MANAGEMENT = 20125
Global Const ERROR_FPX_OBJECT_CREATION_FAILED = 20126
Global Const ERROR_FPX_EXTENSION_FAILED = 20127
Global Const ERROR_FPX_FREE_NULL_PTR = 20128
Global Const ERROR_FPX_INVALID_TILE = 20129
Global Const ERROR_FPX_FILE_IN_USE = 20130
Global Const ERROR_FPX_FILE_CREATE_ERROR = 20131
Global Const ERROR_FPX_FILE_NOT_OPEN_ERROR = 20132
Global Const ERROR_FPX_USER_ABORT = 20133
Global Const ERROR_FPX_OLE_FILE_ERROR = 20134
Global Const ERROR_BAD_TAG = 20140
Global Const ERROR_INVALID_STAMP_SIZE = 20141
Global Const ERROR_BAD_STAMP = 20142
Global Const ERROR_DOCUMENT_NOT_ENABLED = 20143
Global Const ERROR_IMAGE_EMPTY = 20144
Global Const ERROR_NO_CHANGE = 20145
Global Const ERROR_LZW_LOCKED = 20146
Global Const ERROR_FPXEXTENSIONS_LOCKED = 20147
Global Const ERROR_ANN_LOCKED = 20148
Global Const ERROR_DLG_CANCELED = 20150
Global Const ERROR_DLG_FAILED = 20151
Global Const ERROR_ISIS_NOCURSCANNER = 20160
Global Const ERROR_ISIS_SCANDRIVER_NOT_LOADED = 20161
Global Const ERROR_ISIS_CANCEL = 20162
Global Const ERROR_ISIS_BAD_TAG_OR_VALUE = 20163
Global Const ERROR_ISIS_NOT_READY = 20164
Global Const ERROR_ISIS_NO_PAGE = 20165
Global Const ERROR_ISIS_JAM = 20166
Global Const ERROR_ISIS_SCANNER_ERROR = 20167
Global Const ERROR_ISIS_BUSY = 20168
Global Const ERROR_ISIS_FILE_ERROR = 20169
Global Const ERROR_ISIS_NETWORK_ERROR = 20170
Global Const ERROR_ISIS_NOT_INSTALLED = 20171
Global Const ERROR_ISIS_NO_PIXDFLT = 20172
Global Const ERROR_ISIS_PIXVERSION = 20173
Global Const ERROR_ISIS_PERM_NOACCESS = 20174

Global Const ERROR_DOC_NOT_INITIALIZED = 20180
Global Const ERROR_DOC_HANDLE = 20181
Global Const ERROR_DOC_EMPTY = 20182
Global Const ERROR_DOC_INVALID_FONT = 20183
Global Const ERROR_DOC_INVALID_PAGE = 20184
Global Const ERROR_DOC_INVALID_RULE = 20185
Global Const ERROR_DOC_INVALID_ZONE = 20186
Global Const ERROR_DOC_TYPE_ZONE = 20187
Global Const ERROR_DOC_INVALID_COLUMN = 20188
Global Const ERROR_DOC_INVALID_LINE = 20189
Global Const ERROR_DOC_INVALID_WORD = 20190
Global Const ERROR_OCR_LOCKED = 20191
Global Const ERROR_OCR_NOT_INITIALIZED = 20192
Global Const ERROR_OCR_MAX_REGIONS = 20193
Global Const ERROR_OCR_OPTION = 20194
Global Const ERROR_OCR_CONVERT_DIB = 20195
Global Const ERROR_OCR_CANCELED = 20196
Global Const ERROR_OCR_INVALID_OUTPUT = 20197
Global Const ERROR_OCR_BLOCKED = 20198
Global Const ERROR_OCR_RPCMEM = 20199
Global Const ERROR_OCR_FATAL = 20200
Global Const ERROR_OCR_BADTAG = 20201
Global Const ERROR_OCR_BADVAL = 20202
Global Const ERROR_OCR_BADTYPE = 20203
Global Const ERROR_OCR_NOFILE = 20204
Global Const ERROR_OCR_BADTOK = 20205
Global Const ERROR_OCR_BADFMT = 20206
Global Const ERROR_OCR_BADMATCH = 20207
Global Const ERROR_OCR_NOSUPPORT = 20208
Global Const ERROR_OCR_BADID = 20209
Global Const ERROR_OCR_NOLANG = 20210
Global Const ERROR_OCR_LANGOVFL = 20211
Global Const ERROR_OCR_NOISRC = 20212
Global Const ERROR_OCR_NOTIDL = 20213
Global Const ERROR_OCR_NOVER = 20214
Global Const ERROR_OCR_NODRAW = 20215
Global Const ERROR_OCR_MEMERR = 20216
Global Const ERROR_OCR_BADRGN = 20217
Global Const ERROR_OCR_NOICR = 20218
Global Const ERROR_OCR_NOACTV = 20219
Global Const ERROR_OCR_NOMORE = 20220
Global Const ERROR_OCR_NOTWAIT = 20221
Global Const ERROR_OCR_LEXOVFL = 20222
Global Const ERROR_OCR_PREPROC = 20223
Global Const ERROR_OCR_BADFILE = 20224
Global Const ERROR_OCR_BADSCAN = 20225
Global Const ERROR_OCR_NOIMG = 20226
Global Const ERROR_OCR_NOLICN = 20227
Global Const ERROR_OCR_NOLCSRV = 20228
Global Const ERROR_OCR_LMEMERR = 20229
Global Const ERROR_OCR_RESCHNG = 20230
Global Const ERROR_OCR_BADPLGN = 20231
Global Const ERROR_OCR_NOSINK = 20232
Global Const ERROR_OCR_NOSRC = 20233
Global Const ERROR_OCR_NOTOK = 20234
Global Const ERROR_OCR_IMBUFOVFL = 20235
Global Const ERROR_OCR_TMOUT = 20236
Global Const ERROR_OCR_BADVRS = 20237
Global Const ERROR_OCR_TAGNNW = 20238
Global Const ERROR_OCR_SRVCAN = 20239
Global Const ERROR_OCR_WRFAIL = 20240
Global Const ERROR_OCR_SCNCAN = 20241
Global Const ERROR_OCR_RGOCCLD = 20242
Global Const ERROR_OCR_NOTORNT = 20243
Global Const ERROR_OCR_ACCDEN = 20244
Global Const ERROR_OCR_BADUOR = 20245

Global Const ERROR_RECORDING = 20250
Global Const ERROR_COMPRESSOR = 20251
Global Const ERROR_SOUND_DEVICE = 20252
Global Const ERROR_DEVICE_INUSE = 20253
Global Const ERROR_INV_TRACKTYPE = 20254
Global Const ERROR_NO_SOUNDCARD = 20255
Global Const ERROR_NOT_RECORDING = 20256
Global Const ERROR_INV_MODE = 20257
Global Const ERROR_NO_VIDEO_MODULE = 20258
Global Const ERROR_QUEUE_FULL = 20259

Global Const ERROR_HOST_RESOLVE = 20270
Global Const ERROR_CANT_INITIALIZE = 20271
Global Const ERROR_NO_CONNECTION = 20272
Global Const ERROR_HOST_NOT_FOUND = 20273
Global Const ERROR_NOT_SERVER = 20274
Global Const ERROR_NO_CONNECTIONS = 20275
Global Const ERROR_CONNECT_REFUSED = 20276
Global Const ERROR_IS_CONNECTED = 20277
Global Const ERROR_NET_UNREACH = 20278
Global Const ERROR_TIME_OUT = 20279
Global Const ERROR_NET_DOWN = 20280
Global Const ERROR_NO_BUFFERS = 20281
Global Const ERROR_NO_FILE_DESCR = 20282
Global Const ERROR_DATA_QUEUED = 20283
Global Const ERROR_UNKNOWN = 20284
Global Const ERROR_CONNECT_RESET = 20285
Global Const ERROR_TRANSFER_ABORTED = 20286

Global Const ERROR_DSHOW_FAILURE = 20287

Global Const ERROR_REGISTRY_READ = 20288
Global Const ERROR_WAVE_FORMAT = 20289
Global Const ERROR_INSUFICIENT_BUFFER = 20290
Global Const ERROR_WAVE_CONVERT = 20291
Global Const ERROR_MULTIMEDIA_NOT_ENABLED = 20292

Global Const ERROR_CAP_CONNECT = 20293
Global Const ERROR_CAP_DISCONNECT = 20294
Global Const ERROR_DISK_ISFULL = 20295
Global Const ERROR_CAP_OVERLAY = 20296
Global Const ERROR_CAP_PREVIEW = 20297
Global Const ERROR_CAP_COPY = 20298
Global Const ERROR_CAP_WINDOW = 20299
Global Const ERROR_CAP_ISCAPTURING = 20300
Global Const ERROR_NO_STREAMS = 20301
Global Const ERROR_CREATE_STREAM = 20302
Global Const ERROR_FRAME_DELETE = 20303

Global Const ERROR_DXF_FILTER_MISSING = 20309
Global Const ERROR_PAGE_NOT_FOUND = 20310
Global Const ERROR_DELETE_LAST_PAGE = 20311
Global Const ERROR_NO_HOTKEY = 20312
Global Const ERROR_CANNOT_CREATE_HOTKEY_WINDOW = 20313
Global Const ERROR_MEDICAL_NOT_ENABLED = 20314
Global Const ERROR_JBIG_NOT_ENABLED = 20315
Global Const ERROR_UNDO_STACK_EMPTY = 20316
Global Const ERROR_NO_TOOLBAR = 20317
Global Const ERROR_MEDICAL_NET_NOT_ENABLED = 20318
Global Const ERROR_JBIG_FILTER_MISSING = 20319

Global Const ERROR_CAPTURE_STILL_IN_PROCESS = 20320
Global Const ERROR_INVALID_DELAY = 20321
Global Const ERROR_INVALID_COUNT = 20322
Global Const ERROR_INVALID_INTERVAL = 20323
Global Const ERROR_HOTKEY_CONFILCTS_WITH_CANCELKEY = 20324
Global Const ERROR_CAPTURE_INVALID_AREA_TYPE = 20325
Global Const ERROR_CAPTURE_NO_OPTION_STRUCTURE = 20326
Global Const ERROR_CAPTURE_INVALID_FILL_PATTERN = 20327
Global Const ERROR_CAPTURE_INVALID_LINE_STYLE = 20328
Global Const ERROR_CAPTURE_INVALID_INFOWND_POS = 20329
Global Const ERROR_CAPTURE_INVALID_INFOWND_SIZE = 20330
Global Const ERROR_CAPTURE_ZERO_AREA_SIZE = 20331
Global Const ERROR_CAPTURE_FILE_ACCESS_FAILED = 20332
Global Const ERROR_CAPTURE_INVALID_32BIT_EXE_OR_DLL = 20333
Global Const ERROR_CAPTURE_INVALID_RESOURCE_TYPE = 20334
Global Const ERROR_CAPTURE_INVALID_RESOURCE_INDEX = 20335
Global Const ERROR_CAPTURE_NO_ACTIVE_WINDOW = 20336
Global Const ERROR_CAPTURE_CANNOT_CAPTURE_WINDOW = 20337
Global Const ERROR_CAPTURE_STRING_ID_NOT_DEFINED = 20338
Global Const ERROR_CAPTURE_DELAY_LESS_THAN_ZERO = 20339
Global Const ERROR_CAPTURE_NO_MENU = 20340
Global Const ERROR_BROWSE_FAILED = 20350
Global Const ERROR_NOTHING_TO_DO = 20351
Global Const ERROR_INTERNET_NOT_ENABLED = 20352
Global Const L_ERROR_LVKRN_MISSING = 20353

Global Const ERROR_VECTOR_NOT_ENABLED = 20400
Global Const ERROR_VECTOR_DXF_NOT_ENABLED = 20401
Global Const ERROR_VECTOR_DWG_NOT_ENABLED = 20402
Global Const ERROR_VECTOR_MISC_NOT_ENABLED = 20403
Global Const ERROR_TAG_MISSING = 20404
Global Const ERROR_VECTOR_DWF_NOT_ENABLED = 20405
Global Const ERROR_NO_UNDO_STACK = 20406
Global Const ERROR_UNDO_DISABLED = 20407
Global Const ERROR_PDF_NOT_ENABLED = 20408

Global Const ERROR_BARCODE_DIGIT_CHECK = 20410
Global Const ERROR_BARCODE_INVALID_TYPE = 20411
Global Const ERROR_BARCODE_TEXTOUT = 20412
Global Const ERROR_BARCODE_WIDTH = 20413
Global Const ERROR_BARCODE_HEIGHT = 20414
Global Const ERROR_BARCODE_TOSMALL = 20415
Global Const ERROR_BARCODE_STRING = 20416
Global Const ERROR_BARCODE_NOTFOUND = 20417
Global Const ERROR_BARCODE_UNITS = 20418
Global Const ERROR_BARCODE_MULTIPLEMAXCOUNT = 20419
Global Const ERROR_BARCODE_GROUP = 20420
Global Const ERROR_BARCODE_NO_DATA = 20421
Global Const ERROR_BARCODE_NOTFOUND_DUPLICATED = 20422
Global Const ERROR_BARCODE_LAST_DUPLICATED = 20423
Global Const ERROR_BARCODE_STRING_LENGTH = 20424
Global Const ERROR_BARCODE_LOCATION = 20425
Global Const ERROR_BARCODE_1D_LOCKED = 20426
Global Const ERROR_BARCODE_2D_READ_LOCKED = 20427
Global Const ERROR_BARCODE_2D_WRITE_LOCKED = 20428
Global Const ERROR_BARCODE_PDF_READ_LOCKED = 20429
Global Const ERROR_BARCODE_PDF_WRITE_LOCKED = 20430

Global Const ERROR_NET_FIRST = 20435
Global Const ERROR_NET_OUT_OF_HANDLES = 20435
Global Const ERROR_NET_TIMEOUT = 20436
Global Const ERROR_NET_EXTENDED_ERROR = 20437
Global Const ERROR_NET_INTERNAL_ERROR = 20438
Global Const ERROR_NET_INVALID_URL = 20439
Global Const ERROR_NET_UNRECOGNIZED_SCHEME = 20440
Global Const ERROR_NET_NAME_NOT_RESOLVED = 20441
Global Const ERROR_NET_PROTOCOL_NOT_FOUND = 20442
Global Const ERROR_NET_INVALID_OPTION = 20443
Global Const ERROR_NET_BAD_OPTION_LENGTH = 20444
Global Const ERROR_NET_OPTION_NOT_SETTABLE = 20445
Global Const ERROR_NET_SHUTDOWN = 20446
Global Const ERROR_NET_INCORRECT_USER_NAME = 20447
Global Const ERROR_NET_INCORRECT_PASSWORD = 20448
Global Const ERROR_NET_LOGIN_FAILURE = 20449
Global Const ERROR_NET_INVALID_OPERATION = 20450
Global Const ERROR_NET_OPERATION_CANCELLED = 20451
Global Const ERROR_NET_INCORRECT_HANDLE_TYPE = 20452
Global Const ERROR_NET_INCORRECT_HANDLE_STATE = 20453
Global Const ERROR_NET_NOT_PROXY_REQUEST = 20454
Global Const ERROR_NET_REGISTRY_VALUE_NOT_FOUND = 20455
Global Const ERROR_NET_BAD_REGISTRY_PARAMETER = 20456
Global Const ERROR_NET_NO_DIRECT_ACCESS = 20457
Global Const ERROR_NET_NO_CONTEXT = 20458
Global Const ERROR_NET_NO_CALLBACK = 20459
Global Const ERROR_NET_REQUEST_PENDING = 20460
Global Const ERROR_NET_INCORRECT_FORMAT = 20461
Global Const ERROR_NET_ITEM_NOT_FOUND = 20462
Global Const ERROR_NET_CANNOT_CONNECT = 20463
Global Const ERROR_NET_CONNECTION_ABORTED = 20464
Global Const ERROR_NET_CONNECTION_RESET = 20465
Global Const ERROR_NET_FORCE_RETRY = 20466
Global Const ERROR_NET_INVALID_PROXY_REQUEST = 20467
Global Const ERROR_NET_NEED_UI = 20468

Global Const ERROR_NET_HANDLE_EXISTS = 20469
Global Const ERROR_NET_SEC_CERT_DATE_INVALID = 20470
Global Const ERROR_NET_SEC_CERT_CN_INVALID = 20471
Global Const ERROR_NET_HTTP_TO_HTTPS_ON_REDIR = 20472
Global Const ERROR_NET_HTTPS_TO_HTTP_ON_REDIR = 20473
Global Const ERROR_NET_MIXED_SECURITY = 20474
Global Const ERROR_NET_CHG_POST_IS_NON_SECURE = 20475
Global Const ERROR_NET_POST_IS_NON_SECURE = 20476
Global Const ERROR_NET_CLIENT_AUTH_CERT_NEEDED = 20477
Global Const ERROR_NET_INVALID_CA = 20478
Global Const ERROR_NET_CLIENT_AUTH_NOT_SETUP = 20479
Global Const ERROR_NET_ASYNC_THREAD_FAILED = 20480
Global Const ERROR_NET_REDIRECT_SCHEME_CHANGE = 20481
Global Const ERROR_NET_DIALOG_PENDING = 20482
Global Const ERROR_NET_RETRY_DIALOG = 20483
Global Const ERROR_NET_HTTPS_HTTP_SUBMIT_REDIR = 20484
Global Const ERROR_NET_INSERT_CDROM = 20485

Global Const ERROR_NET_HTTP_HEADER_NOT_FOUND = 20486
Global Const ERROR_NET_HTTP_DOWNLEVEL_SERVER = 20487
Global Const ERROR_NET_HTTP_INVALID_SERVER_RESPONSE = 20488
Global Const ERROR_NET_HTTP_INVALID_HEADER = 20489
Global Const ERROR_NET_HTTP_INVALID_QUERY_REQUEST = 20490
Global Const ERROR_NET_HTTP_HEADER_ALREADY_EXISTS = 20491
Global Const ERROR_NET_HTTP_REDIRECT_FAILED = 20492
Global Const ERROR_NET_HTTP_NOT_REDIRECTED = 20493
Global Const ERROR_NET_HTTP_COOKIE_NEEDS_CONFIRMATION = 20494
Global Const ERROR_NET_HTTP_COOKIE_DECLINED = 20495
Global Const ERROR_NET_HTTP_REDIRECT_NEEDS_CONFIRMATION = 20496

Global Const ERROR_NET_NO_OPEN_REQUEST = 20497

Global Const ERROR_VECTOR_IS_LOCKED = 20500
Global Const ERROR_VECTOR_IS_EMPTY = 20501
Global Const ERROR_VECTOR_LAYER_NOT_FOUND = 20502
Global Const ERROR_VECTOR_LAYER_IS_LOCKED = 20503
Global Const ERROR_VECTOR_LAYER_ALREADY_EXISTS = 20504
Global Const ERROR_VECTOR_OBJECT_NOT_FOUND = 20505
Global Const ERROR_VECTOR_INVALID_OBJECT_TYPE = 20506
Global Const ERROR_VECTOR_PEN_NOT_FOUND = 20507
Global Const ERROR_VECTOR_BRUSH_NOT_FOUND = 20508
Global Const ERROR_VECTOR_FONT_NOT_FOUND = 20509
Global Const ERROR_VECTOR_BITMAP_NOT_FOUND = 20510
Global Const ERROR_VECTOR_POINT_NOT_FOUND = 20511
Global Const ERROR_VECTOR_ENGINE_NOT_FOUND = 20512
Global Const ERROR_VECTOR_INVALID_ENGINE = 20513
Global Const ERROR_VECTOR_CLIPBOARD = 20514
Global Const ERROR_VECTOR_CLIPBOARD_IS_EMPTY = 20515
Global Const ERROR_VECTOR_CANT_ADD_TEXT = 20516
Global Const ERROR_VECTOR_CANT_READ_WMF = 20517
Global Const ERROR_VECTOR_GROUP_NOT_FOUND = 20518
Global Const ERROR_VECTOR_GROUP_ALREADY_EXISTS = 20519

Global Const ERROR_JP2_FAILURE = 20530
Global Const ERROR_JP2_SIGNATURE = 20531
Global Const ERROR_JP2_UNSUPPORTED = 20532
Global Const ERROR_J2K_FAILURE = 20533
Global Const ERROR_J2K_NO_SOC = 20534
Global Const ERROR_J2K_NO_SOT = 20535
Global Const ERROR_J2K_INFORMATION_SET = 20536
Global Const ERROR_J2K_LOW_TARGET_SIZE = 20537
Global Const ERROR_J2K_DECOMPOSITION_LEVEL = 20538
Global Const ERROR_J2K_MARKER_VALUE = 20539
Global Const ERROR_J2K_UNSUPPORTED = 20540
Global Const ERROR_J2K_FILTER_MISSING = 20541
Global Const ERROR_J2K_LOCKED = 20542

Global Const ERROR_TWAIN_NO_LIBRARY = 20560
Global Const ERROR_TWAIN_NOT_AVAILABLE = 20560
Global Const ERROR_TWAIN_INVALID_DLL = 20561
Global Const ERROR_TWAIN_NOT_INITIALIZED = 20562
Global Const ERROR_TWAIN_CANCELED = 20563
Global Const ERROR_TWAIN_CHECK_STATUS = 20564
Global Const ERROR_TWAIN_END_OF_LIST = 20565
Global Const ERROR_TWAIN_CAP_NOT_SUPPORTED = 20566
Global Const ERROR_TWAIN_SOURCE_NOT_OPEN = 20567
Global Const ERROR_TWAIN_BAD_VALUE = 20568
Global Const ERROR_TWAIN_INVALID_STATE = 20569
Global Const ERROR_TWAIN_CAPS_NEG_NOT_ENDED = 20570
Global Const ERROR_TWAIN_OPEN_FILE = 20571
Global Const ERROR_TWAIN_INV_HANDLE = 20572
Global Const ERROR_TWAIN_WRITE_TO_FILE = 20573
Global Const ERROR_TWAIN_INV_VERSION_NUM = 20574
Global Const ERROR_TWAIN_READ_FROM_FILE = 20575
Global Const ERROR_TWAIN_NOT_VALID_FILE = 20576
Global Const ERROR_TWAIN_INV_ACCESS_RIGHT = 20577

Global Const ERROR_PAINT_INTERNAL = 20600
Global Const ERROR_PAINT_INV_DATA = 20601
Global Const ERROR_PAINT_NO_RESOURCES = 20602
Global Const ERROR_PAINT_NOT_ENABLED = 20603

Global Const ERROR_CONTAINER_INV_HANDLE = 20630
Global Const ERROR_CONTAINER_INV_OPERATION = 20631
Global Const ERROR_CONTAINER_NO_RESOURCES = 20632

Global Const ERROR_TOOLBAR_NO_RESOURCES = 20660
Global Const ERROR_TOOLBAR_INV_STATE = 20661
Global Const ERROR_TOOLBAR_INV_HANDLE = 20662

Global Const ERROR_AUTOMATION_INV_HANDLE = 20690
Global Const ERROR_AUTOMATION_INV_STATE = 20691



Global Const RESIZE_TYPE = 0
Global Const RESIZEREGION_TYPE = 1

Global Const PAINT_AUTO = 0
Global Const PAINT_FIXED = 1
Global Const PAINT_NETSCAPE = 2

Global Const SO_LEAD = 0
Global Const SO_JFIF = 1
Global Const SO_JTIF = 2
Global Const SO_AWD = 3
Global Const SO_CALS = 4
Global Const SO_CUR = 5
Global Const SO_CCITT = 6
Global Const SO_DIC_GRAY = 7
Global Const SO_DIC_COLOR = 8
Global Const SO_EXIF = 9
Global Const SO_FAX = 10
Global Const SO_EPS = 11
Global Const SO_FPX = 12
Global Const SO_GEM = 13
Global Const SO_GIF = 14
Global Const SO_ICO = 15
Global Const SO_IOCA = 16
Global Const SO_PCT = 17
Global Const SO_MAC = 18
Global Const SO_MSP = 19
Global Const SO_OS2 = 20
Global Const SO_PCX = 21
Global Const SO_PNG = 22
Global Const SO_PSD = 23
Global Const SO_RAS = 24
Global Const SO_TGA = 25
Global Const SO_TIF = 26
Global Const SO_WBMP = 27
Global Const SO_WBMP_RLE = 28
Global Const SO_WFX = 29
Global Const SO_WMF = 30
Global Const SO_WPG = 31
Global Const NUM_SAVE_TYPES = 32
Global Const NUM_OPEN_TYPES = 25

Global Const QF_CUSTOM = 9

Global gDitheringType As Integer
Global gBitonalScaling As Integer
Global gPaintScaling As Integer
Global gPalette As Integer
Global gUseNetscape As Boolean
Global gNumChildren As Integer
Global gInProcess As Integer
Global gEndApp As Integer
Global gLoadRepaint As Boolean

Public Const BITSPIXEL = 12
Public Const PLANES = 14

Global Const IDM_TOOLNONE = 0
Global Const IDM_TOOLRECT = 1
Global Const IDM_TOOLELLIPSE = 2
Global Const IDM_TOOLRNDRECT = 3
Global Const IDM_TOOLFREEHAND = 4

Global Kids() As Variant

#If Win32 Then
    Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
#Else
    Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
    Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
    Declare Function GetDC Lib "User" (ByVal hWnd As Integer) As Integer
    Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hdc As Integer) As Integer
#End If

Global Const UNITS_PER_INCH = 1000

'Stores iVal (1000ths of inches or pixels) in string szVal
Public Function UnitToString(ByVal iVal As Long, ByVal bInches As Boolean) As String
    On Error GoTo ErrorHandler
    Dim szVal As String
    Dim dVal As Double
   
    If (bInches = True) Then
        dVal = iVal / UNITS_PER_INCH
        szVal = CStr(dVal)
    Else 'pixels
        dVal = iVal
        szVal = CStr(dVal)
    End If
    UnitToString = szVal
    Exit Function
ErrorHandler:
    MsgBox "Error in LeadTools.UnitToString.", vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
End Function

'Converts szVal to 1000ths of inches or pixels
Public Function StringToUnit(iVal As Long, szVal As String, ByVal bInches As Boolean) As Boolean
    Dim dVal As Double
    Dim bRet As Boolean

    bRet = False
On Error GoTo STRTOUNITERR
    dVal = CDbl(szVal)
    bRet = True
On Error GoTo 0
    If (bInches = True) Then
        iVal = dVal * UNITS_PER_INCH
    Else
        iVal = dVal
    End If
STRTOUNITERR:
    StringToUnit = bRet
End Function

Public Function InchesToPixels(ByVal iInches As Long, ByVal iRes As Long) As Long
    On Error GoTo ErrorHandler
    Dim iPixels As Long

    iPixels = iInches * iRes / UNITS_PER_INCH
    InchesToPixels = iPixels
    Exit Function
ErrorHandler:
    MsgBox "Error in LeadTools.InchesToPixels.", vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
End Function

Public Function PixelsToInches(ByVal iPixels As Long, ByVal iRes As Long) As Long
    On Error GoTo ErrorHandler
    Dim iInches As Long

    iInches = iPixels * UNITS_PER_INCH / iRes
    PixelsToInches = iInches
    Exit Function
ErrorHandler:
    MsgBox "Error in LeadTools.PixelsToInches.", vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
End Function


