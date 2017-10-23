﻿// -- FILE ------------------------------------------------------------------
// name       : IRtfVisualBreak.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.22
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace RtfPipe
{

	// ------------------------------------------------------------------------
	public interface IRtfVisualBreak : IRtfVisual
	{

		// ----------------------------------------------------------------------
		RtfVisualBreakKind BreakKind { get; }

	} // interface IRtfVisualBreak

} // namespace RtfPipe
// -- EOF -------------------------------------------------------------------
