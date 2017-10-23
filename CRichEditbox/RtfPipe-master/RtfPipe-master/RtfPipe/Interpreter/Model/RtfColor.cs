﻿// -- FILE ------------------------------------------------------------------
// name       : RtfColor.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using RtfPipe.Sys;

namespace RtfPipe.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfColor : IRtfColor
	{

		// ----------------------------------------------------------------------
		public static readonly IRtfColor Black = new RtfColor( 0, 0, 0 );
		public static readonly IRtfColor White = new RtfColor( 255, 255, 255 );

		// ----------------------------------------------------------------------
		public RtfColor( int red, int green, int blue )
		{
			if ( red < 0 || red > 255 )
			{
				throw new RtfColorException( Strings.InvalidColor( red ) );
			}
			if ( green < 0 || green > 255 )
			{
				throw new RtfColorException( Strings.InvalidColor( green ) );
			}
			if ( blue < 0 || blue > 255 )
			{
				throw new RtfColorException( Strings.InvalidColor( blue ) );
			}
			this.red = red;
			this.green = green;
			this.blue = blue;
		} // RtfColor

		// ----------------------------------------------------------------------
		public int Red
		{
			get { return red; }
		} // Red

		// ----------------------------------------------------------------------
		public int Green
		{
			get { return green; }
		} // Green

		// ----------------------------------------------------------------------
		public int Blue
		{
			get { return blue; }
		} // Blue

		// ----------------------------------------------------------------------
		public override bool Equals( object obj )
		{
			if ( obj == this )
			{
				return true;
			}
			
			if ( obj == null || GetType() != obj.GetType() )
			{
				return false;
			}

			return IsEqual( obj );
		} // Equals

		// ----------------------------------------------------------------------
		public override int GetHashCode()
		{
			return HashTool.AddHashCode( GetType().GetHashCode(), ComputeHashCode() );
		} // GetHashCode

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return "Color{" + red + "," + green + "," + blue + "}";
		} // ToString

		// ----------------------------------------------------------------------
		private bool IsEqual( object obj )
		{
			RtfColor compare = obj as RtfColor; // guaranteed to be non-null
			return compare != null && red == compare.red &&
				green == compare.green &&
				blue == compare.blue;
		} // IsEqual

		// ----------------------------------------------------------------------
		private int ComputeHashCode()
		{
			int hash = red;
			hash = HashTool.AddHashCode( hash, green );
			hash = HashTool.AddHashCode( hash, blue );
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		// members
		private readonly int red;
		private readonly int green;
		private readonly int blue;

	} // class RtfColor

} // namespace RtfPipe.Model
// -- EOF -------------------------------------------------------------------
