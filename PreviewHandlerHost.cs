﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using System.IO;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;


namespace DocumentsModule
{
    class PreviewHandlerHost
    {
        private object mCurrentPreviewHandler;
        private Guid mCurrentPreviewHandlerGUID;
        private Stream mCurrentPreviewHandlerStream;
        private string mErrorMessage;
        private Color mBorderColor;
        private BorderStyle mBorderStyle;
        /*private string ErrorMessage
        {
            get { return mErrorMessage; }
            set
            {
                mErrorMessage = value;
                Invalidate();	// repaint the control
            }
        }*/
        /// <summary>
        /// Gets the GUID of the current preview handler.
        /// </summary>
        [Browsable(false), ReadOnly(true)]
        public Guid CurrentPreviewHandler
        {
            get
            {
                return mCurrentPreviewHandlerGUID;
            }
        }
        /// <summary>
        /// Gets or sets the background colour of this PreviewHandlerHost.
        /// </summary>
        [DefaultValue(typeof(Color), "White")]
        /*public override System.Drawing.Color BackColor
        {
            get
            {
                return base.BackColor;
            }
            set
            {
                base.BackColor = value;
            }
        }*/
        /// <summary>
        /// Gets or sets the colour of the border to draw around the control.
        /// </summary>
        //[DefaultValue(typeof(Color), "ControlDark")]
        /*public Color BorderColor
        {
            get
            {
                return mBorderColor;
            }
            set
            {
                mBorderColor = value;
                Invalidate();
            }
        }*/
        /// <summary>
        /// Gets or sets the style of the border to draw around the control.
        /// </summary>
        //[DefaultValue(BorderStyle.FixedSingle)]
        /*public BorderStyle BorderStyle
        {
            get
            {
                return mBorderStyle;
            }
            set
            {
                mBorderStyle = value;
                Invalidate();
            }
        }*/

        /// <summary>
        /// Initialialises a new instance of the PreviewHandlerHost class.
        /// </summary>
        /*public PreviewHandlerHost() : base()
        {
            mCurrentPreviewHandlerGUID = Guid.Empty;
            BackColor = Color.White;
            mBorderColor = SystemColors.ControlDark;
            mBorderStyle = BorderStyle.FixedSingle;
            Size = new Size(320, 240);

            // display default error message (no file)
            ErrorMessage = "No file loaded.";

            // enable transparency
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.UserPaint, true);
        }*/

        /// <summary>
        /// Releases the unmanaged resources used by the PreviewHandlerHost and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing"></param>
        /*protected override void Dispose(bool disposing)
        {
            UnloadPreviewHandler();

            if (mCurrentPreviewHandler != null)
            {
                Marshal.FinalReleaseComObject(mCurrentPreviewHandler);
                mCurrentPreviewHandler = null;
                GC.Collect();
            }

            base.Dispose(disposing);
        }*/

        /// <summary>
        /// Returns the GUID of the preview handler associated with the specified file.
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private Guid GetPreviewHandlerGUID(string filename)
        {
            // open the registry key corresponding to the file extension
            RegistryKey ext = Registry.ClassesRoot.OpenSubKey(Path.GetExtension(filename));
            if (ext != null)
            {
                // open the key that indicates the GUID of the preview handler type
                RegistryKey test = ext.OpenSubKey("shellex\\{8895b1c6-b41f-4c1c-a562-0d564250836f}");
                if (test != null) return new Guid(Convert.ToString(test.GetValue(null)));

                // sometimes preview handlers are declared on key for the class
                string className = Convert.ToString(ext.GetValue(null));
                if (className != null)
                {
                    test = Registry.ClassesRoot.OpenSubKey(className + "\\shellex\\{8895b1c6-b41f-4c1c-a562-0d564250836f}");
                    if (test != null) return new Guid(Convert.ToString(test.GetValue(null)));
                }
            }

            return Guid.Empty;
        }

        /// <summary>
        /// Paints the error message text on the PreviewHandlerHost control.
        /// </summary>
        /// <param name="e"></param>
        /*protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (mErrorMessage != String.Empty)
            {
                // paint the error message
                TextRenderer.DrawText(
                    e.Graphics,
                    mErrorMessage,
                    Font,
                    ClientRectangle,
                    ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis
                );
            }

            // border
            if (mBorderStyle == BorderStyle.Fixed3D)
            {
                ControlPaint.DrawBorder(e.Graphics, ClientRectangle, mBorderColor, ButtonBorderStyle.Inset);
            }
            else if (mBorderStyle == BorderStyle.FixedSingle)
            {
                using (Pen pen = new Pen(mBorderColor))
                {
                    Rectangle rect = ClientRectangle;
                    rect.Width--;
                    rect.Height--;
                    e.Graphics.DrawRectangle(pen, rect);
                }
            }
        }*/

        /// <summary>
        /// Resizes the hosted preview handler when this PreviewHandlerHost is resized.
        /// </summary>
        /// <param name="e"></param>
        /*protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            if (mCurrentPreviewHandler is IPreviewHandler)
            {
                // update the preview handler's bounds to match the control's
                Rectangle r = ClientRectangle;
                ((IPreviewHandler)mCurrentPreviewHandler).SetRect(ref r);
            }
        }*/

        /// <summary>
        /// Opens the specified file using the appropriate preview handler and displays the result in this PreviewHandlerHost.
        /// </summary>
        /// <param name=""></param>
        /// <returns></rfilenameeturns>
        public bool Open(string filename)
        {
            UnloadPreviewHandler();

            if (String.IsNullOrEmpty(filename))
            {
                //ErrorMessage = "No file loaded.";
                return false;
            }

            // try to get GUID for the preview handler
            Guid guid = GetPreviewHandlerGUID(filename);
            //ErrorMessage = "";

            if (guid != Guid.Empty)
            {
                try
                {
                    if (guid != mCurrentPreviewHandlerGUID)
                    {
                        mCurrentPreviewHandlerGUID = guid;

                        // need to instantiate a different COM type (file format has changed)
                        if (mCurrentPreviewHandler != null) Marshal.FinalReleaseComObject(mCurrentPreviewHandler);

                        // use reflection to instantiate the preview handler type
                        Type comType = Type.GetTypeFromCLSID(mCurrentPreviewHandlerGUID);
                        mCurrentPreviewHandler = Activator.CreateInstance(comType);
                    }

                    if (mCurrentPreviewHandler is IInitializeWithFile)
                    {
                        // some handlers accept a filename
                        ((IInitializeWithFile)mCurrentPreviewHandler).Initialize(filename, 0);
                    }
                    else if (mCurrentPreviewHandler is IInitializeWithStream)
                    {
                        if (File.Exists(filename))
                        {
                            // other handlers want an IStream (in this case, a file stream)
                            mCurrentPreviewHandlerStream = File.Open(filename, FileMode.Open);
                            StreamWrapper stream = new StreamWrapper(mCurrentPreviewHandlerStream);
                            ((IInitializeWithStream)mCurrentPreviewHandler).Initialize(stream, 0);
                        }
                        else
                        {
                            //ErrorMessage = "File not found.";
                        }
                    }
                    else if (mCurrentPreviewHandler is IInitializeWithItem)
                    {
                        // a third category exists, must be initialised with a shell item
                        IShellItem shellItem;
                        //SHCreateItemFromParsingName(filename, IntPtr.Zero, new Guid(GUID_ISHELLITEM), out shellItem);
                        //((IInitializeWithItem)mCurrentPreviewHandler).Initialize(shellItem, 0);
                    }

                    if (mCurrentPreviewHandler is IPreviewHandler)
                    {
                        // bind the preview handler to the control's bounds and preview the content
                        //Rectangle r = ClientRectangle;
                        //((IPreviewHandler)mCurrentPreviewHandler).SetWindow(Handle, ref r);
                        ((IPreviewHandler)mCurrentPreviewHandler).DoPreview();

                        return true;
                    }
                }
                catch (Exception ex)
                {
                    //ErrorMessage = "Preview could not be generated.\n" + ex.Message;
                }
            }
            else
            {
                //ErrorMessage = "No preview available.";
            }

            return false;
        }

        /// <summary>
        /// Opens the specified stream using the preview handler COM type with the provided GUID and displays the result in this PreviewHandlerHost.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="previewHandler"></param>
        /// <returns></returns>
        /*public bool Open(Stream stream, Guid previewHandler)
        {
            UnloadPreviewHandler();

            if (stream == null)
            {
                ErrorMessage = "No file loaded.";
                return false;
            }

            ErrorMessage = "";

            if (previewHandler != Guid.Empty)
            {
                try
                {
                    if (previewHandler != mCurrentPreviewHandlerGUID)
                    {
                        mCurrentPreviewHandlerGUID = previewHandler;

                        // need to instantiate a different COM type (file format has changed)
                        if (mCurrentPreviewHandler != null) Marshal.FinalReleaseComObject(mCurrentPreviewHandler);

                        // use reflection to instantiate the preview handler type
                        Type comType = Type.GetTypeFromCLSID(mCurrentPreviewHandlerGUID);
                        mCurrentPreviewHandler = Activator.CreateInstance(comType);
                    }

                    if (mCurrentPreviewHandler is IInitializeWithStream)
                    {
                        // must wrap the stream to provide compatibility with IStream
                        mCurrentPreviewHandlerStream = stream;
                        StreamWrapper wrapped = new StreamWrapper(mCurrentPreviewHandlerStream);
                        ((IInitializeWithStream)mCurrentPreviewHandler).Initialize(wrapped, 0);
                    }

                    if (mCurrentPreviewHandler is IPreviewHandler)
                    {
                        // bind the preview handler to the control's bounds and preview the content
                        Rectangle r = ClientRectangle;
                        ((IPreviewHandler)mCurrentPreviewHandler).SetWindow(Handle, ref r);
                        ((IPreviewHandler)mCurrentPreviewHandler).DoPreview();

                        return true;
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessage = "Preview could not be generated.\n" + ex.Message;
                }
            }
            else
            {
                ErrorMessage = "No preview available.";
            }

            return false;
        }*/

        /// <summary>
        /// Unloads the preview handler hosted in this PreviewHandlerHost and closes the file stream.
        /// </summary>
        public void UnloadPreviewHandler()
        {
            if (mCurrentPreviewHandler is IPreviewHandler)
            {
                // explicitly unload the content
                ((IPreviewHandler)mCurrentPreviewHandler).Unload();
            }
            if (mCurrentPreviewHandlerStream != null)
            {
                mCurrentPreviewHandlerStream.Close();
                mCurrentPreviewHandlerStream = null;
            }
        }

        #region P/Invoke

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void SHCreateItemFromParsingName(
            [In][MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            [In] IntPtr pbc, [In][MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [Out][MarshalAs(UnmanagedType.Interface, IidParameterIndex = 2)] out IShellItem ppv
        );

        #endregion
    }

    #region COM Interop

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("8895b1c6-b41f-4c1c-a562-0d564250836f")]
    internal interface IPreviewHandler
    {
        void SetWindow(IntPtr hwnd, ref Rectangle rect);
        void SetRect(ref Rectangle rect);
        void DoPreview();
        void Unload();
        void SetFocus();
        void QueryFocus(out IntPtr phwnd);
        [PreserveSig]
        uint TranslateAccelerator(ref Message pmsg);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("b7d14566-0509-4cce-a71f-0a554233bd9b")]
    internal interface IInitializeWithFile
    {
        void Initialize([MarshalAs(UnmanagedType.LPWStr)] string pszFilePath, uint grfMode);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
    internal interface IInitializeWithStream
    {
        void Initialize(IStream pstream, uint grfMode);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("7F73BE3F-FB79-493C-A6C7-7EE14E245841")]
    interface IInitializeWithItem
    {
        void Initialize(IShellItem psi, uint grfMode);
    }

    //[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid(PreviewHandlerHost.GUID_ISHELLITEM)]
    interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)]Guid bhid, [MarshalAs(UnmanagedType.LPStruct)]Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    };

    /// <summary>
    /// Provides a bare-bones implementation of System.Runtime.InteropServices.IStream that wraps an System.IO.Stream.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal class StreamWrapper : IStream
    {

        private System.IO.Stream mInner;

        /// <summary>
        /// Initialises a new instance of the StreamWrapper class, using the specified System.IO.Stream.
        /// </summary>
        /// <param name="inner"></param>
        public StreamWrapper(System.IO.Stream inner)
        {
            mInner = inner;
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        /// <param name="ppstm"></param>
        public void Clone(out IStream ppstm)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        /// <param name="grfCommitFlags"></param>
        public void Commit(int grfCommitFlags)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        /// <param name="pstm"></param>
        /// <param name="cb"></param>
        /// <param name="pcbRead"></param>
        /// <param name="pcbWritten"></param>
        public void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        /// <param name="libOffset"></param>
        /// <param name="cb"></param>
        /// <param name="dwLockType"></param>
        public void LockRegion(long libOffset, long cb, int dwLockType)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Reads a sequence of bytes from the underlying System.IO.Stream.
        /// </summary>
        /// <param name="pv"></param>
        /// <param name="cb"></param>
        /// <param name="pcbRead"></param>
        public void Read(byte[] pv, int cb, IntPtr pcbRead)
        {
            long bytesRead = mInner.Read(pv, 0, cb);
            if (pcbRead != IntPtr.Zero) Marshal.WriteInt64(pcbRead, bytesRead);
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        public void Revert()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Advances the stream to the specified position.
        /// </summary>
        /// <param name="dlibMove"></param>
        /// <param name="dwOrigin"></param>
        /// <param name="plibNewPosition"></param>
        public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
        {
            long pos = mInner.Seek(dlibMove, (System.IO.SeekOrigin)dwOrigin);
            if (plibNewPosition != IntPtr.Zero) Marshal.WriteInt64(plibNewPosition, pos);
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        /// <param name="libNewSize"></param>
        public void SetSize(long libNewSize)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Returns details about the stream, including its length, type and name.
        /// </summary>
        /// <param name="pstatstg"></param>
        /// <param name="grfStatFlag"></param>
        public void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag)
        {
            pstatstg = new System.Runtime.InteropServices.ComTypes.STATSTG();
            pstatstg.cbSize = mInner.Length;
            pstatstg.type = 2; // stream type
            pstatstg.pwcsName = (mInner is FileStream) ? ((FileStream)mInner).Name : String.Empty;
        }

        /// <summary>
        /// This operation is not supported.
        /// </summary>
        /// <param name="libOffset"></param>
        /// <param name="cb"></param>
        /// <param name="dwLockType"></param>
        public void UnlockRegion(long libOffset, long cb, int dwLockType)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Writes a sequence of bytes to the underlying System.IO.Stream.
        /// </summary>
        /// <param name="pv"></param>
        /// <param name="cb"></param>
        /// <param name="pcbWritten"></param>
        public void Write(byte[] pv, int cb, IntPtr pcbWritten)
        {
            mInner.Write(pv, 0, cb);
            if (pcbWritten != IntPtr.Zero) Marshal.WriteInt64(pcbWritten, (Int64)cb);
        }
    }

    #endregion

}