using System;
using System.Collections;
using System.Collections.Generic;

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

            public As GetExceptionText            {
                strExceptionstring  ;
                ;
                switch ExceptionCode;
                Case EXCEPTION_ACCESS_VIOLATION:         strExceptionstring = "Access Violation";
                Case EXCEPTION_DATATYPE_MISALIGNMENT:    strExceptionstring = "Data Type Misalignment";
                Case EXCEPTION_BREAKPOINT:               strExceptionstring = "Breakpoint";
                Case EXCEPTION_SINGLE_STEP:              strExceptionstring = "Single Step";
                Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED:    strExceptionstring = "Array Bounds Exceeded";
                Case EXCEPTION_FLT_DENORMAL_OPERAND:     strExceptionstring = "Float Denormal Operand";
                Case EXCEPTION_FLT_DIVIDE_BY_ZERO:       strExceptionstring = "Divide By Zero";
                Case EXCEPTION_FLT_INEXACT_RESULT:       strExceptionstring = "Floating Point Inexact Result";
                Case EXCEPTION_FLT_INVALID_OPERATION:    strExceptionstring = "Invalid Operation";
                Case EXCEPTION_FLT_OVERFLOW:             strExceptionstring = "Float Overflow";
                Case EXCEPTION_FLT_STACK_CHECK:          strExceptionstring = "Float Stack Check";
                Case EXCEPTION_FLT_UNDERFLOW:            strExceptionstring = "Float Underflow";
                Case EXCEPTION_INT_DIVIDE_BY_ZERO:       strExceptionstring = "Integer Divide By Zero";
                Case EXCEPTION_INT_OVERFLOW:             strExceptionstring = "Integer Overflow";
                Case EXCEPTION_PRIVILEGED_INSTRUCTION:   strExceptionstring = "Privileged Instruction";
                Case EXCEPTION_IN_PAGE_ERROR:            strExceptionstring = "In Page Error";
                Case EXCEPTION_ILLEGAL_INSTRUCTION:      strExceptionstring = "Illegal Instruction";
                Case EXCEPTION_NONCONTINUABLE_EXCEPTION: strExceptionstring = "Non Continuable Exception";
                Case EXCEPTION_STACK_OVERFLOW:           strExceptionstring = "Stack Overflow";
                Case EXCEPTION_INVALID_DISPOSITION:      strExceptionstring = "Invalid Disposition";
                Case EXCEPTION_GUARD_PAGE_VIOLATION:     strExceptionstring = "Guard Page Violation";
                Case EXCEPTION_INVALID_HANDLE:           strExceptionstring = "Invalid Handle";
                Case EXCEPTION_CONTROL_C_EXIT:           strExceptionstring = "Control-C Exit";
                Case }
            else
"00000000" + Hex(ExceptionCode).Substring("00000000" + Hex(ExceptionCode).Length - 8) + ")";
                };
                GetExceptionText = strExceptionstring;
            }

            public As MyExceptionFilter            {
                Dim  ;
                Dim  ;
                ;
                // ' Get the current exception record.;
                Rec = ExceptionPtrs.pExceptionRecord;
                ;
                // ' If Rec.pExceptionRecord is not zero, then it is a nested exception and;
                // ' Rec.pExceptionRecord points to another EXCEPTION_RECORD structure.  Follow;
                // ' the pointers back to the original exception.;
                Do Until Rec.pExceptionRecord = 0;
                //    ' A friendly declaration of CopyMemory.;
                CopyExceptionRecord Rec, Rec.pExceptionRecord, Len(Rec);
                Loop;
                ;
                // ' Translate the exception code into a user-friendly string.;
                strException = GetExceptionText(Rec.ExceptionCode);
                ;
                // ' Raise an error to return control to the calling procedure.;
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
