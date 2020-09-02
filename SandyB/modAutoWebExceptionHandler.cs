using System;

namespace MetX.SliceAndDice
{
		class modExceptionHandler
		{

			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;
			public = Const;

			public As GetExceptionText			{
				String strExceptionString;

				switch ExceptionCode;
				Case EXCEPTION_ACCESS_VIOLATION:         strExceptionString = "Access Violation";
				Case EXCEPTION_DATATYPE_MISALIGNMENT:    strExceptionString = "Data Type Misalignment";
				Case EXCEPTION_BREAKPOINT:               strExceptionString = "Breakpoint";
				Case EXCEPTION_SINGLE_STEP:              strExceptionString = "Single Step";
				Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED:    strExceptionString = "Array Bounds Exceeded";
				Case EXCEPTION_FLT_DENORMAL_OPERAND:     strExceptionString = "Float Denormal Operand";
				Case EXCEPTION_FLT_DIVIDE_BY_ZERO:       strExceptionString = "Divide By Zero";
				Case EXCEPTION_FLT_INEXACT_RESULT:       strExceptionString = "Floating Point Inexact Result";
				Case EXCEPTION_FLT_INVALID_OPERATION:    strExceptionString = "Invalid Operation";
				Case EXCEPTION_FLT_OVERFLOW:             strExceptionString = "Float Overflow";
				Case EXCEPTION_FLT_STACK_CHECK:          strExceptionString = "Float Stack Check";
				Case EXCEPTION_FLT_UNDERFLOW:            strExceptionString = "Float Underflow";
				Case EXCEPTION_INT_DIVIDE_BY_ZERO:       strExceptionString = "Integer Divide By Zero";
				Case EXCEPTION_INT_OVERFLOW:             strExceptionString = "Integer Overflow";
				Case EXCEPTION_PRIVILEGED_INSTRUCTION:   strExceptionString = "Privileged Instruction";
				Case EXCEPTION_IN_PAGE_ERROR:            strExceptionString = "In Page Error";
				Case EXCEPTION_ILLEGAL_INSTRUCTION:      strExceptionString = "Illegal Instruction";
				Case EXCEPTION_NONCONTINUABLE_EXCEPTION: strExceptionString = "Non Continuable Exception";
				Case EXCEPTION_STACK_OVERFLOW:           strExceptionString = "Stack Overflow";
				Case EXCEPTION_INVALID_DISPOSITION:      strExceptionString = "Invalid Disposition";
				Case EXCEPTION_GUARD_PAGE_VIOLATION:     strExceptionString = "Guard Page Violation";
				Case EXCEPTION_INVALID_HANDLE:           strExceptionString = "Invalid Handle";
				Case EXCEPTION_CONTROL_C_EXIT:           strExceptionString = "Control-C Exit";
				Case }
				else
				{:                               strExceptionString = "Unknown (+H" + Right$("00000000" + Hex(ExceptionCode), 8) + ")";
				};
				GetExceptionText = strExceptionString;
			}
			public As MyExceptionFilter			{
				EXCEPTION_RECORD Rec;
				String strException;

				//  Get the current exception record.;
				Rec = ExceptionPtrs.pExceptionRecord;

				//  If Rec.pExceptionRecord is not zero, then it is a nested exception and;
				//  Rec.pExceptionRecord points to another EXCEPTION_RECORD structure.  Follow;
				//  the pointers back to the original exception.;
				Do Until Rec.pExceptionRecord = 0;
				//  A friendly declaration of CopyMemory.;
				CopyExceptionRecord Rec, Rec.pExceptionRecord, Len(Rec);
				Loop;

				//  Translate the exception code into a user-friendly string.;
				strException = GetExceptionText(Rec.ExceptionCode);

				//  Raise an error to return control to the calling procedure.;
				Err.Raise 10000, "MyExceptionFilter", strException;
			}
			public void ExceptionStartup()
			{
				Call SetUnhandledExceptionFilter(AddressOf MyExceptionFilter);
			}
			public void ExceptionShutdown()
			{
				Call SetUnhandledExceptionFilter(0);
			}
		}
}
