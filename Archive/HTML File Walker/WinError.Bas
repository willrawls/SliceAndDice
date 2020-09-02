Attribute VB_Name = "MWinError"
' The configuration registry database operation completed successfully.
Public Const ERROR_SUCCESS = 0&

'   Incorrect function.
Public Const ERROR_INVALID_FUNCTION = 1 '  dderror

'   The system cannot find the file specified.
Public Const ERROR_FILE_NOT_FOUND = 2&

'   The system cannot find the path specified.
Public Const ERROR_PATH_NOT_FOUND = 3&

'   The system cannot open the file.
Public Const ERROR_TOO_MANY_OPEN_FILES = 4&

'   Access is denied.
Public Const ERROR_ACCESS_DENIED = 5&

'   The handle is invalid.
Public Const ERROR_INVALID_HANDLE = 6&

'   The storage control blocks were destroyed.
Public Const ERROR_ARENA_TRASHED = 7&

'   Not enough storage is available to process this command.
Public Const ERROR_NOT_ENOUGH_MEMORY = 8 '  dderror

'   The storage control block address is invalid.
Public Const ERROR_INVALID_BLOCK = 9&

'   The environment is incorrect.
Public Const ERROR_BAD_ENVIRONMENT = 10&

'   An attempt was made to load a program with an
'   incorrect format.
Public Const ERROR_BAD_FORMAT = 11&

'   The access code is invalid.
Public Const ERROR_INVALID_ACCESS = 12&

'   The data is invalid.
Public Const ERROR_INVALID_DATA = 13&

'   Not enough storage is available to complete this operation.
Public Const ERROR_OUTOFMEMORY = 14&

'   The system cannot find the drive specified.
Public Const ERROR_INVALID_DRIVE = 15&

'   The directory cannot be removed.
Public Const ERROR_CURRENT_DIRECTORY = 16&

'   The system cannot move the file
'   to a different disk drive.
Public Const ERROR_NOT_SAME_DEVICE = 17&

'   There are no more files.
Public Const ERROR_NO_MORE_FILES = 18&

'   The media is write protected.
Public Const ERROR_WRITE_PROTECT = 19&

'   The system cannot find the device specified.
Public Const ERROR_BAD_UNIT = 20&

'   The device is not ready.
Public Const ERROR_NOT_READY = 21&

'   The device does not recognize the command.
Public Const ERROR_BAD_COMMAND = 22&

'   Data error (cyclic redundancy check)
Public Const ERROR_CRC = 23&

'   The program issued a command but the
'   command length is incorrect.
Public Const ERROR_BAD_LENGTH = 24&

'   The drive cannot locate a specific
'   area or track on the disk.
Public Const ERROR_SEEK = 25&

'   The specified disk or diskette cannot be accessed.
Public Const ERROR_NOT_DOS_DISK = 26&

'   The drive cannot find the sector requested.
Public Const ERROR_SECTOR_NOT_FOUND = 27&

'   The printer is out of paper.
Public Const ERROR_OUT_OF_PAPER = 28&

'   The system cannot write to the specified device.
Public Const ERROR_WRITE_FAULT = 29&

'   The system cannot read from the specified device.
Public Const ERROR_READ_FAULT = 30&

'   A device attached to the system is not functioning.
Public Const ERROR_GEN_FAILURE = 31&

'   The process cannot access the file because
'   it is being used by another process.
Public Const ERROR_SHARING_VIOLATION = 32&

'   The process cannot access the file because
'   another process has locked a portion of the file.
Public Const ERROR_LOCK_VIOLATION = 33&

'   The wrong diskette is in the drive.
'   Insert %2 (Volume Serial Number: %3)
'   into drive %1.
Public Const ERROR_WRONG_DISK = 34&

'   Too many files opened for sharing.
Public Const ERROR_SHARING_BUFFER_EXCEEDED = 36&

'   Reached end of file.
Public Const ERROR_HANDLE_EOF = 38&

'   The disk is full.
Public Const ERROR_HANDLE_DISK_FULL = 39&

'   The network request is not supported.
Public Const ERROR_NOT_SUPPORTED = 50&

'   The remote computer is not available.
Public Const ERROR_REM_NOT_LIST = 51&

'   A duplicate name exists on the network.
Public Const ERROR_DUP_NAME = 52&

'   The network path was not found.
Public Const ERROR_BAD_NETPATH = 53&

'   The network is busy.
Public Const ERROR_NETWORK_BUSY = 54&

'   The specified network resource or device is no longer
'   available.
Public Const ERROR_DEV_NOT_EXIST = 55 '  dderror

'   The network BIOS command limit has been reached.
Public Const ERROR_TOO_MANY_CMDS = 56&

'   A network adapter hardware error occurred.
Public Const ERROR_ADAP_HDW_ERR = 57&

'   The specified server cannot perform the requested
'   operation.
Public Const ERROR_BAD_NET_RESP = 58&

'   An unexpected network error occurred.
Public Const ERROR_UNEXP_NET_ERR = 59&

'   The remote adapter is not compatible.
Public Const ERROR_BAD_REM_ADAP = 60&

'   The printer queue is full.
Public Const ERROR_PRINTQ_FULL = 61&

'   Space to store the file waiting to be printed is
'   not available on the server.
Public Const ERROR_NO_SPOOL_SPACE = 62&

'   Your file waiting to be printed was deleted.
Public Const ERROR_PRINT_CANCELLED = 63&

'   The specified network name is no longer available.
Public Const ERROR_NETNAME_DELETED = 64&

'   Network access is denied.
Public Const ERROR_NETWORK_ACCESS_DENIED = 65&

'   The network resource type is not correct.
Public Const ERROR_BAD_DEV_TYPE = 66&

'   The network name cannot be found.
Public Const ERROR_BAD_NET_NAME = 67&

'   The name limit for the local computer network
'   adapter card was exceeded.
Public Const ERROR_TOO_MANY_NAMES = 68&

'   The network BIOS session limit was exceeded.
Public Const ERROR_TOO_MANY_SESS = 69&

'   The remote server has been paused or is in the
'   process of being started.
Public Const ERROR_SHARING_PAUSED = 70&

'   The network request was not accepted.
Public Const ERROR_REQ_NOT_ACCEP = 71&

'   The specified printer or disk device has been paused.
Public Const ERROR_REDIR_PAUSED = 72&

'   The file exists.
Public Const ERROR_FILE_EXISTS = 80&

'   The directory or file cannot be created.
Public Const ERROR_CANNOT_MAKE = 82&

'   Fail on INT 24
Public Const ERROR_FAIL_I24 = 83&

'   Storage to process this request is not available.
Public Const ERROR_OUT_OF_STRUCTURES = 84&

'   The local device name is already in use.
Public Const ERROR_ALREADY_ASSIGNED = 85&

'   The specified network password is not correct.
Public Const ERROR_INVALID_PASSWORD = 86&

'   The parameter is incorrect.
Public Const ERROR_INVALID_PARAMETER = 87 '  dderror

'   A write fault occurred on the network.
Public Const ERROR_NET_WRITE_FAULT = 88&

'   The system cannot start another process at
'   this ti
Public Const ERROR_NO_PROC_SLOTS = 89&

'   Cannot create another system semaphore.
Public Const ERROR_TOO_MANY_SEMAPHORES = 100&

'   The exclusive semaphore is owned by another process.
Public Const ERROR_EXCL_SEM_ALREADY_OWNED = 101&

'   The semaphore is set and cannot be closed.
Public Const ERROR_SEM_IS_SET = 102&

'   The semaphore cannot be set again.
Public Const ERROR_TOO_MANY_SEM_REQUESTS = 103&

'   Cannot request exclusive semaphores at interrupt ti
Public Const ERROR_INVALID_AT_INTERRUPT_TIME = 104&

'   The previous ownership of this semaphore has ended.
Public Const ERROR_SEM_OWNER_DIED = 105&

'   Insert the diskette for drive %1.
Public Const ERROR_SEM_USER_LIMIT = 106&

'   Program stopped because alternate diskette was not inserted.
Public Const ERROR_DISK_CHANGE = 107&

'   The disk is in use or locked by
'   another process.
Public Const ERROR_DRIVE_LOCKED = 108&

'   The pipe has been ended.
Public Const ERROR_BROKEN_PIPE = 109&

'   The system cannot open the
'   device or file specified.
Public Const ERROR_OPEN_FAILED = 110&

'   The file name is too long.
Public Const ERROR_BUFFER_OVERFLOW = 111&

'   There is not enough space on the disk.
Public Const ERROR_DISK_FULL = 112&

'   No more internal file identifiers available.
Public Const ERROR_NO_MORE_SEARCH_HANDLES = 113&

'   The target internal file identifier is incorrect.
Public Const ERROR_INVALID_TARGET_HANDLE = 114&

'   The IOCTL call made by the application program is
'   not correct.
Public Const ERROR_INVALID_CATEGORY = 117&

'   The verify-on-write switch parameter value is not
'   correct.
Public Const ERROR_INVALID_VERIFY_SWITCH = 118&

'   The system does not support the command requested.
Public Const ERROR_BAD_DRIVER_LEVEL = 119&

'   This function is only valid in Windows NT mode.
Public Const ERROR_CALL_NOT_IMPLEMENTED = 120&

'   The semaphore timeout period has expired.
Public Const ERROR_SEM_TIMEOUT = 121&

'   The data area passed to a system call is too
'   small.
Public Const ERROR_INSUFFICIENT_BUFFER = 122 '  dderror

'   The filename, directory name, or volume label syntax is incorrect.
Public Const ERROR_INVALID_NAME = 123&

'   The system call level is not correct.
Public Const ERROR_INVALID_LEVEL = 124&

'   The disk has no volume label.
Public Const ERROR_NO_VOLUME_LABEL = 125&

'   The specified module could not be found.
Public Const ERROR_MOD_NOT_FOUND = 126&

'   The specified procedure could not be found.
Public Const ERROR_PROC_NOT_FOUND = 127&

'   There are no child processes to wait for.
Public Const ERROR_WAIT_NO_CHILDREN = 128&

'   The %1 application cannot be run in Windows NT mode.
Public Const ERROR_CHILD_NOT_COMPLETE = 129&

'   Attempt to use a file handle to an open disk partition for an
'   operation other than raw disk I/O.
Public Const ERROR_DIRECT_ACCESS_HANDLE = 130&

'   An attempt was made to move the file pointer before the beginning of the file.
Public Const ERROR_NEGATIVE_SEEK = 131&

'   The file pointer cannot be set on the specified device or file.
Public Const ERROR_SEEK_ON_DEVICE = 132&

'   A JOIN or SUBST command
'   cannot be used for a drive that
'   contains previously joined drives.
Public Const ERROR_IS_JOIN_TARGET = 133&

'   An attempt was made to use a
'   JOIN or SUBST command on a drive that has
'   already been joined.
Public Const ERROR_IS_JOINED = 134&

'   An attempt was made to use a
'   JOIN or SUBST command on a drive that has
'   already been substituted.
Public Const ERROR_IS_SUBSTED = 135&

'   The system tried to delete
'   the JOIN of a drive that is not joined.
Public Const ERROR_NOT_JOINED = 136&

'   The system tried to delete the
'   substitution of a drive that is not substituted.
Public Const ERROR_NOT_SUBSTED = 137&

'   The system tried to join a drive
'   to a directory on a joined drive.
Public Const ERROR_JOIN_TO_JOIN = 138&

'   The system tried to substitute a
'   drive to a directory on a substituted drive.
Public Const ERROR_SUBST_TO_SUBST = 139&

'   The system tried to join a drive to
'   a directory on a substituted drive.
Public Const ERROR_JOIN_TO_SUBST = 140&

'   The system tried to SUBST a drive
'   to a directory on a joined drive.
Public Const ERROR_SUBST_TO_JOIN = 141&

'   The system cannot perform a JOIN or SUBST at this ti
Public Const ERROR_BUSY_DRIVE = 142&

'   The system cannot join or substitute a
'   drive to or for a directory on the same drive.
Public Const ERROR_SAME_DRIVE = 143&

'   The directory is not a subdirectory of the root directory.
Public Const ERROR_DIR_NOT_ROOT = 144&

'   The directory is not empty.
Public Const ERROR_DIR_NOT_EMPTY = 145&

'   The path specified is being used in
'   a substitute.
Public Const ERROR_IS_SUBST_PATH = 146&

'   Not enough resources are available to
'   process this command.
Public Const ERROR_IS_JOIN_PATH = 147&

'   The path specified cannot be used at this ti
Public Const ERROR_PATH_BUSY = 148&

'   An attempt was made to join
'   or substitute a drive for which a directory
'   on the drive is the target of a previous
'   substitute.
Public Const ERROR_IS_SUBST_TARGET = 149&

'   System trace information was not specified in your
'   CONFIG.SYS file, or tracing is disallowed.
Public Const ERROR_SYSTEM_TRACE = 150&

'   The number of specified semaphore events for
'   DosMuxSemWait is not correct.
Public Const ERROR_INVALID_EVENT_COUNT = 151&

'   DosMuxSemWait did not execute; too many semaphores
'   are already set.
Public Const ERROR_TOO_MANY_MUXWAITERS = 152&

'   The DosMuxSemWait list is not correct.
Public Const ERROR_INVALID_LIST_FORMAT = 153&

'   The volume label you entered exceeds the
'   11 character limit.  The first 11 characters were written
'   to disk.  Any characters that exceeded the 11 character limit
'   were automatically deleted.
Public Const ERROR_LABEL_TOO_LONG = 154&

'   Cannot create another thread.
Public Const ERROR_TOO_MANY_TCBS = 155&

'   The recipient process has refused the signal.
Public Const ERROR_SIGNAL_REFUSED = 156&

'   The segment is already discarded and cannot be locked.
Public Const ERROR_DISCARDED = 157&

'   The segment is already unlocked.
Public Const ERROR_NOT_LOCKED = 158&

'   The address for the thread ID is not correct.
Public Const ERROR_BAD_THREADID_ADDR = 159&

'   The argument string passed to DosExecPgm is not correct.
Public Const ERROR_BAD_ARGUMENTS = 160&

'   The specified path is invalid.
Public Const ERROR_BAD_PATHNAME = 161&

'   A signal is already pending.
Public Const ERROR_SIGNAL_PENDING = 162&

'   No more threads can be created in the system.
Public Const ERROR_MAX_THRDS_REACHED = 164&

'   Unable to lock a region of a file.
Public Const ERROR_LOCK_FAILED = 167&

'   The requested resource is in use.
Public Const ERROR_BUSY = 170&

'   A lock request was not outstanding for the supplied cancel region.
Public Const ERROR_CANCEL_VIOLATION = 173&

'   The file system does not support atomic changes to the lock type.
Public Const ERROR_ATOMIC_LOCKS_NOT_SUPPORTED = 174&

'   The system detected a segment number that was not correct.
Public Const ERROR_INVALID_SEGMENT_NUMBER = 180&

'   The operating system cannot run %1.
Public Const ERROR_INVALID_ORDINAL = 182&

'   Cannot create a file when that file already exists.
Public Const ERROR_ALREADY_EXISTS = 183&

'   The flag passed is not correct.
Public Const ERROR_INVALID_FLAG_NUMBER = 186&

'   The specified system semaphore name was not found.
Public Const ERROR_SEM_NOT_FOUND = 187&

'   The operating system cannot run %1.
Public Const ERROR_INVALID_STARTING_CODESEG = 188&

'   The operating system cannot run %1.
Public Const ERROR_INVALID_STACKSEG = 189&

'   The operating system cannot run %1.
Public Const ERROR_INVALID_MODULETYPE = 190&

'   Cannot run %1 in Windows NT mode.
Public Const ERROR_INVALID_EXE_SIGNATURE = 191&

'   The operating system cannot run %1.
Public Const ERROR_EXE_MARKED_INVALID = 192&

'   %1 is not a valid Windows NT application.
Public Const ERROR_BAD_EXE_FORMAT = 193&

'   The operating system cannot run %1.
Public Const ERROR_ITERATED_DATA_EXCEEDS_64k = 194&

'   The operating system cannot run %1.
Public Const ERROR_INVALID_MINALLOCSIZE = 195&

'   The operating system cannot run this
'   application program.
Public Const ERROR_DYNLINK_FROM_INVALID_RING = 196&

'   The operating system is not presently
'   configured to run this application.
Public Const ERROR_IOPL_NOT_ENABLED = 197&

'   The operating system cannot run %1.
Public Const ERROR_INVALID_SEGDPL = 198&

'   The operating system cannot run this
'   application program.
Public Const ERROR_AUTODATASEG_EXCEEDS_64k = 199&

'   The code segment cannot be greater than or equal to 64KB.
Public Const ERROR_RING2SEG_MUST_BE_MOVABLE = 200&

'   The operating system cannot run %1.
Public Const ERROR_RELOC_CHAIN_XEEDS_SEGLIM = 201&

'   The operating system cannot run %1.
Public Const ERROR_INFLOOP_IN_RELOC_CHAIN = 202&

'   The system could not find the environment
'   option that was entered.
Public Const ERROR_ENVVAR_NOT_FOUND = 203&

'   No process in the command subtree has a
'   signal handler.
Public Const ERROR_NO_SIGNAL_SENT = 205&

'   The filename or extension is too long.
Public Const ERROR_FILENAME_EXCED_RANGE = 206&

'   The ring 2 stack is in use.
Public Const ERROR_RING2_STACK_IN_USE = 207&

'   The Global filename characters,  or ?, are entered
'   incorrectly or too many Global filename characters are specified.
Public Const ERROR_META_EXPANSION_TOO_LONG = 208&

'   The signal being posted is not correct.
Public Const ERROR_INVALID_SIGNAL_NUMBER = 209&

'   The signal handler cannot be set.
Public Const ERROR_THREAD_1_INACTIVE = 210&

'   The segment is locked and cannot be reallocated.
Public Const ERROR_LOCKED = 212&

'   Too many dynamic link modules are attached to this
'   program or dynamic link module.
Public Const ERROR_TOO_MANY_MODULES = 214&

'   Can't nest calls to LoadModule.
Public Const ERROR_NESTING_NOT_ALLOWED = 215&

'   The pipe state is invalid.
Public Const ERROR_BAD_PIPE = 230&

'   All pipe instances are busy.
Public Const ERROR_PIPE_BUSY = 231&

'   The pipe is being closed.
Public Const ERROR_NO_DATA = 232&

'   No process is on the other end of the pipe.
Public Const ERROR_PIPE_NOT_CONNECTED = 233&

'   More data is available.
Public Const ERROR_MORE_DATA = 234 '  dderror

'   The session was cancelled.
Public Const ERROR_VC_DISCONNECTED = 240&

'   The specified extended attribute name was invalid.
Public Const ERROR_INVALID_EA_NAME = 254&

'   The extended attributes are inconsistent.
Public Const ERROR_EA_LIST_INCONSISTENT = 255&

'   No more data is available.
Public Const ERROR_NO_MORE_ITEMS = 259&

'   The Copy API cannot be used.
Public Const ERROR_CANNOT_COPY = 266&

'   The directory name is invalid.
Public Const ERROR_DIRECTORY = 267&

'   The extended attributes did not fit in the buffer.
Public Const ERROR_EAS_DIDNT_FIT = 275&

'   The extended attribute file on the mounted file system is corrupt.
Public Const ERROR_EA_FILE_CORRUPT = 276&

'   The extended attribute table file is full.
Public Const ERROR_EA_TABLE_FULL = 277&

'   The specified extended attribute handle is invalid.
Public Const ERROR_INVALID_EA_HANDLE = 278&

'   The mounted file system does not support extended attributes.
Public Const ERROR_EAS_NOT_SUPPORTED = 282&

'   Attempt to release mutex not owned by caller.
Public Const ERROR_NOT_OWNER = 288&

'   Too many posts were made to a semaphore.
Public Const ERROR_TOO_MANY_POSTS = 298&

'   The system cannot find message for message number 0x%1
'   in message file for %2.
Public Const ERROR_MR_MID_NOT_FOUND = 317&

'   Attempt to access invalid address.
Public Const ERROR_INVALID_ADDRESS = 487&

'   Arithmetic result exceeded 32 bits.
Public Const ERROR_ARITHMETIC_OVERFLOW = 534&

'   There is a process on other end of the pipe.
Public Const ERROR_PIPE_CONNECTED = 535&

'   Waiting for a process to open the other end of the pipe.
Public Const ERROR_PIPE_LISTENING = 536&

'   Access to the extended attribute was denied.
Public Const ERROR_EA_ACCESS_DENIED = 994&

'   The I/O operation has been aborted because of either a thread exit
'   or an application request.
Public Const ERROR_OPERATION_ABORTED = 995&

'   Overlapped I/O event is not in a signalled state.
Public Const ERROR_IO_INCOMPLETE = 996&

'   Overlapped I/O operation is in progress.
Public Const ERROR_IO_PENDING = 997 '  dderror

'   Invalid access to memory location.
Public Const ERROR_NOACCESS = 998&

'   Error performing inpage operation.
Public Const ERROR_SWAPERROR = 999&

'   Recursion too deep, stack overflowed.
Public Const ERROR_STACK_OVERFLOW = 1001&

'   The window cannot act on the sent message.
Public Const ERROR_INVALID_MESSAGE = 1002&

'   Cannot complete this function.
Public Const ERROR_CAN_NOT_COMPLETE = 1003&

'   Invalid flags.
Public Const ERROR_INVALID_FLAGS = 1004&

'   The volume does not contain a recognized file system.
'   Please make sure that all required file system drivers are loaded and that the
'   volume is not corrupt.
Public Const ERROR_UNRECOGNIZED_VOLUME = 1005&

'   The volume for a file has been externally altered such that the
'   opened file is no longer valid.
Public Const ERROR_FILE_INVALID = 1006&

'   The requested operation cannot be performed in full-screen mode.
Public Const ERROR_FULLSCREEN_MODE = 1007&

'   An attempt was made to reference a token that does not exist.
Public Const ERROR_NO_TOKEN = 1008&

'   The configuration registry database is corrupt.
Public Const ERROR_BADDB = 1009&

'   The configuration registry key is invalid.
Public Const ERROR_BADKEY = 1010&

'   The configuration registry key could not be opened.
Public Const ERROR_CANTOPEN = 1011&

'   The configuration registry key could not be read.
Public Const ERROR_CANTREAD = 1012&

'   The configuration registry key could not be written.
Public Const ERROR_CANTWRITE = 1013&

'   One of the files in the Registry database had to be recovered
'   by use of a log or alternate copy.  The recovery was successful.
Public Const ERROR_REGISTRY_RECOVERED = 1014&

'   The Registry is corrupt. The structure of one of the files that contains
'   Registry data is corrupt, or the system's image of the file in memory
'   is corrupt, or the file could not be recovered because the alternate
'   copy or log was absent or corrupt.
Public Const ERROR_REGISTRY_CORRUPT = 1015&

'   An I/O operation initiated by the Registry failed unrecoverably.
'   The Registry could not read in, or write out, or flush, one of the files
'   that contain the system's image of the Registry.
Public Const ERROR_REGISTRY_IO_FAILED = 1016&

'   The system has attempted to load or restore a file into the Registry, but the
'   specified file is not in a Registry file format.
Public Const ERROR_NOT_REGISTRY_FILE = 1017&

'   Illegal operation attempted on a Registry key which has been marked for deletion.
Public Const ERROR_KEY_DELETED = 1018&

'   System could not allocate the required space in a Registry log.
Public Const ERROR_NO_LOG_SPACE = 1019&

'   Cannot create a symbolic link in a Registry key that already
'   has subkeys or values.
Public Const ERROR_KEY_HAS_CHILDREN = 1020&

'   Cannot create a stable subkey under a volatile parent key.
Public Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&

'   A notify change request is being completed and the information
'   is not being returned in the caller's buffer. The caller now
'   needs to enumerate the files to find the changes.
Public Const ERROR_NOTIFY_ENUM_DIR = 1022&

'   A stop control has been sent to a service which other running services
'   are dependent on.
Public Const ERROR_DEPENDENT_SERVICES_RUNNING = 1051&

'   The requested control is not valid for this service
Public Const ERROR_INVALID_SERVICE_CONTROL = 1052&

'   The service did not respond to the start or control request in a timely
'   fashion.
Public Const ERROR_SERVICE_REQUEST_TIMEOUT = 1053&

'   A thread could not be created for the service.
Public Const ERROR_SERVICE_NO_THREAD = 1054&

'   The service database is locked.
Public Const ERROR_SERVICE_DATABASE_LOCKED = 1055&

'   An instance of the service is already running.
Public Const ERROR_SERVICE_ALREADY_RUNNING = 1056&

'   The account name is invalid or does not exist.
Public Const ERROR_INVALID_SERVICE_ACCOUNT = 1057&

'   The specified service is disabled and cannot be started.
Public Const ERROR_SERVICE_DISABLED = 1058&

'   Circular service dependency was specified.
Public Const ERROR_CIRCULAR_DEPENDENCY = 1059&

'   The specified service does not exist as an installed service.
Public Const ERROR_SERVICE_DOES_NOT_EXIST = 1060&

'   The service cannot accept control messages at this ti
Public Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL = 1061&

'   The service has not been started.
Public Const ERROR_SERVICE_NOT_ACTIVE = 1062&

'   The service process could not connect to the service controller.
Public Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT = 1063&

'   An exception occurred in the service when handling the control request.
Public Const ERROR_EXCEPTION_IN_SERVICE = 1064&

'   The database specified does not exist.
Public Const ERROR_DATABASE_DOES_NOT_EXIST = 1065&

'   The service has returned a service-specific error code.
Public Const ERROR_SERVICE_SPECIFIC_ERROR = 1066&

'   The process terminated unexpectedly.
Public Const ERROR_PROCESS_ABORTED = 1067&

'   The dependency service or group failed to start.
Public Const ERROR_SERVICE_DEPENDENCY_FAIL = 1068&

'   The service did not start due to a logon failure.
Public Const ERROR_SERVICE_LOGON_FAILED = 1069&

'   After starting, the service hung in a start-pending state.
Public Const ERROR_SERVICE_START_HANG = 1070&

'   The specified service database lock is invalid.
Public Const ERROR_INVALID_SERVICE_LOCK = 1071&

'   The specified service has been marked for deletion.
Public Const ERROR_SERVICE_MARKED_FOR_DELETE = 1072&

'   The specified service already exists.
Public Const ERROR_SERVICE_EXISTS = 1073&

'   The system is currently running with the last-known-good configuration.
Public Const ERROR_ALREADY_RUNNING_LKG = 1074&

'   The dependency service does not exist or has been marked for
'   deletion.
Public Const ERROR_SERVICE_DEPENDENCY_DELETED = 1075&

'   The current boot has already been accepted for use as the
'   last-known-good control set.
Public Const ERROR_BOOT_ALREADY_ACCEPTED = 1076&

'   No attempts to start the service have been made since the last boot.
Public Const ERROR_SERVICE_NEVER_STARTED = 1077&

'   The name is already in use as either a service name or a service display
'   na
Public Const ERROR_DUPLICATE_SERVICE_NAME = 1078&

'   The physical end of the tape has been reached.
Public Const ERROR_END_OF_MEDIA = 1100&

'   A tape access reached a filemark.
Public Const ERROR_FILEMARK_DETECTED = 1101&

'   Beginning of tape or partition was encountered.
Public Const ERROR_BEGINNING_OF_MEDIA = 1102&

'   A tape access reached the end of a set of files.
Public Const ERROR_SETMARK_DETECTED = 1103&

'   No more data is on the tape.
Public Const ERROR_NO_DATA_DETECTED = 1104&

'   Tape could not be partitioned.
Public Const ERROR_PARTITION_FAILURE = 1105&

'   When accessing a new tape of a multivolume partition, the current
'   blocksize is incorrect.
Public Const ERROR_INVALID_BLOCK_LENGTH = 1106&

'   Tape partition information could not be found when loading a tape.
Public Const ERROR_DEVICE_NOT_PARTITIONED = 1107&

'   Unable to lock the media eject mechanism.
Public Const ERROR_UNABLE_TO_LOCK_MEDIA = 1108&

'   Unable to unload the media.
Public Const ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109&

'   Media in drive may have changed.
Public Const ERROR_MEDIA_CHANGED = 1110&

'   The I/O bus was reset.
Public Const ERROR_BUS_RESET = 1111&

'   No media in drive.
Public Const ERROR_NO_MEDIA_IN_DRIVE = 1112&

'   No mapping for the Unicode character exists in the target multi-byte code page.
Public Const ERROR_NO_UNICODE_TRANSLATION = 1113&

'   A dynamic link library (DLL) initialization routine failed.
Public Const ERROR_DLL_INIT_FAILED = 1114&

'   A system shutdown is in progress.
Public Const ERROR_SHUTDOWN_IN_PROGRESS = 1115&

'   Unable to abort the system shutdown because no shutdown was in progress.
Public Const ERROR_NO_SHUTDOWN_IN_PROGRESS = 1116&

'   The request could not be performed because of an I/O device error.
Public Const ERROR_IO_DEVICE = 1117&

'   No serial device was successfully initialized.  The serial driver will unload.
Public Const ERROR_SERIAL_NO_DEVICE = 1118&

'   Unable to open a device that was sharing an interrupt request (IRQ)
'   with other devices. At least one other device that uses that IRQ
'   was already opened.
Public Const ERROR_IRQ_BUSY = 1119&

'   A serial I/O operation was completed by another write to the serial port.
'   (The IOCTL_SERIAL_XOFF_COUNTER reached zero.)
Public Const ERROR_MORE_WRITES = 1120&

'   A serial I/O operation completed because the time-out period expired.
'   (The IOCTL_SERIAL_XOFF_COUNTER did not reach zero.)
Public Const ERROR_COUNTER_TIMEOUT = 1121&

'   No ID address mark was found on the floppy disk.
Public Const ERROR_FLOPPY_ID_MARK_NOT_FOUND = 1122&

'   Mismatch between the floppy disk sector ID field and the floppy disk
'   controller track address.
Public Const ERROR_FLOPPY_WRONG_CYLINDER = 1123&

'   The floppy disk controller reported an error that is not recognized
'   by the floppy disk driver.
Public Const ERROR_FLOPPY_UNKNOWN_ERROR = 1124&

'   The floppy disk controller returned inconsistent results in its registers.
Public Const ERROR_FLOPPY_BAD_REGISTERS = 1125&

'   While accessing the hard disk, a recalibrate operation failed, even after retries.
Public Const ERROR_DISK_RECALIBRATE_FAILED = 1126&

'   While accessing the hard disk, a disk operation failed even after retries.
Public Const ERROR_DISK_OPERATION_FAILED = 1127&

'   While accessing the hard disk, a disk controller reset was needed, but
'   even that failed.
Public Const ERROR_DISK_RESET_FAILED = 1128&

'   Physical end of tape encountered.
Public Const ERROR_EOM_OVERFLOW = 1129&

'   Not enough server storage is available to process this command.
Public Const ERROR_NOT_ENOUGH_SERVER_MEMORY = 1130&

'   A potential deadlock condition has been detected.
Public Const ERROR_POSSIBLE_DEADLOCK = 1131&

'   The base address or the file offset specified does not have the proper
'   alignment.
Public Const ERROR_MAPPED_ALIGNMENT = 1132&

' NEW for Win32
Public Const ERROR_INVALID_PIXEL_FORMAT = 2000
Public Const ERROR_BAD_DRIVER = 2001
Public Const ERROR_INVALID_WINDOW_STYLE = 2002
Public Const ERROR_METAFILE_NOT_SUPPORTED = 2003
Public Const ERROR_TRANSFORM_NOT_SUPPORTED = 2004
Public Const ERROR_CLIPPING_NOT_SUPPORTED = 2005
Public Const ERROR_UNKNOWN_PRINT_MONITOR = 3000
Public Const ERROR_PRINTER_DRIVER_IN_USE = 3001
Public Const ERROR_SPOOL_FILE_NOT_FOUND = 3002
Public Const ERROR_SPL_NO_STARTDOC = 3003
Public Const ERROR_SPL_NO_ADDJOB = 3004
Public Const ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED = 3005
Public Const ERROR_PRINT_MONITOR_ALREADY_INSTALLED = 3006
Public Const ERROR_WINS_INTERNAL = 4000
Public Const ERROR_CAN_NOT_DEL_LOCAL_WINS = 4001
Public Const ERROR_STATIC_INIT = 4002
Public Const ERROR_INC_BACKUP = 4003
Public Const ERROR_FULL_BACKUP = 4004
Public Const ERROR_REC_NON_EXISTENT = 4005
Public Const ERROR_RPL_NOT_ALLOWED = 4006

Public Const ERROR_SEVERITY_SUCCESS = &H0
Public Const ERROR_SEVERITY_INFORMATIONAL = &H40000000
Public Const ERROR_SEVERITY_WARNING = &H80000000
Public Const ERROR_SEVERITY_ERROR = &HC0000000
