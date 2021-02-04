using System;
using System.Collections.Generic;
using System.Text;
using ExcelKit.Core.Constraint.Enums;
using NPOI.SS.UserModel;

namespace ExcelKit.Core.Constraint.Mappings
{
	/// <summary>
	/// 外部使用颜色与NPOI内置颜色映射
	/// </summary>
	internal class ColorMapping
	{
		/// <summary>
		/// 颜色映射
		/// </summary>
		static Dictionary<DefineColor, short> ColorMappings = new Dictionary<DefineColor, short>()
		{
			{DefineColor.Black, IndexedColors.Black.Index}, {DefineColor.PaleBlue, IndexedColors.PaleBlue.Index},
			{DefineColor.Rose, IndexedColors.Rose.Index}, {DefineColor.Lavender, IndexedColors.Lavender.Index},
			{DefineColor.Tan, IndexedColors.Tan.Index}, {DefineColor.LightBlue, IndexedColors.LightBlue.Index},
			{DefineColor.Aqua, IndexedColors.Aqua.Index}, {DefineColor.Lime, IndexedColors.Lime.Index},
			{DefineColor.Gold, IndexedColors.Gold.Index}, {DefineColor.LightOrange, IndexedColors.LightOrange.Index},
			{DefineColor.Orange, IndexedColors.Orange.Index}, {DefineColor.BlueGrey, IndexedColors.BlueGrey.Index},
			{DefineColor.Grey40Percent, IndexedColors.Grey40Percent.Index}, {DefineColor.DarkTeal, IndexedColors.DarkTeal.Index},
			{DefineColor.SeaGreen, IndexedColors.SeaGreen.Index}, {DefineColor.DarkGreen, IndexedColors.DarkGreen.Index},
			{DefineColor.OliveGreen, IndexedColors.OliveGreen.Index}, {DefineColor.Brown, IndexedColors.Brown.Index},
			{DefineColor.Plum, IndexedColors.Plum.Index}, {DefineColor.Indigo, IndexedColors.Indigo.Index},
			{DefineColor.Grey80Percent, IndexedColors.Grey80Percent.Index}, {DefineColor.Automatic, IndexedColors.Automatic.Index},
			{DefineColor.LightGreen, IndexedColors.LightGreen.Index}, {DefineColor.LightTurquoise, IndexedColors.LightTurquoise.Index},
			{DefineColor.LightYellow, IndexedColors.LightYellow.Index}, {DefineColor.LightCornflowerBlue, IndexedColors.LightCornflowerBlue.Index},
			{DefineColor.White, IndexedColors.White.Index}, {DefineColor.Red, IndexedColors.Red.Index},
			{DefineColor.BrightGreen, IndexedColors.BrightGreen.Index}, {DefineColor.Blue, IndexedColors.Blue.Index},
			{DefineColor.Yellow, IndexedColors.Yellow.Index}, {DefineColor.Pink, IndexedColors.Pink.Index},
			{DefineColor.Turquoise, IndexedColors.Turquoise.Index}, {DefineColor.DarkRed, IndexedColors.DarkRed.Index},
			{DefineColor.Green, IndexedColors.Green.Index}, {DefineColor.SkyBlue, IndexedColors.SkyBlue.Index},
			{DefineColor.DarkYellow, IndexedColors.DarkYellow.Index}, {DefineColor.DarkBlue, IndexedColors.DarkBlue.Index},
			{DefineColor.Teal, IndexedColors.Teal.Index}, {DefineColor.Grey25Percent, IndexedColors.Grey25Percent.Index},
			{DefineColor.Grey50Percent, IndexedColors.Grey50Percent.Index}, {DefineColor.CornflowerBlue, IndexedColors.CornflowerBlue.Index},
			{DefineColor.Maroon, IndexedColors.Maroon.Index}, {DefineColor.LemonChiffon, IndexedColors.LemonChiffon.Index},
			{DefineColor.Orchid, IndexedColors.Orchid.Index}, {DefineColor.Coral, IndexedColors.Coral.Index},
			{DefineColor.RoyalBlue, IndexedColors.RoyalBlue.Index}, {DefineColor.Violet, IndexedColors.Violet.Index},
		};

		/// <summary>
		/// 获取颜色索引（映射不上时默认返回黑色）
		/// </summary>
		/// <param name="color"></param>
		/// <returns></returns>
		public static short GetColorIndex(DefineColor color)
		{
			return ColorMappings.ContainsKey(color) ? ColorMappings[color] : IndexedColors.Black.Index;
		}
	}
}
