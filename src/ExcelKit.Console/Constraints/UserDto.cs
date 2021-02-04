using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using DocumentFormat.OpenXml.Drawing;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Constraint.Enums;
using ExcelKit.Core.Infrastructure.Converter;

namespace ExcelKit.Consoles
{
	public class UserBaseDto
	{
		[ExcelKit(Desc = "编码", Width = 20, Align = TextAlign.Right, FontColor = DefineColor.Red)]
		public string Code { get; set; } = "20210123";
	}

	public class UserDto : UserBaseDto
	{
		[ExcelKit(Desc = "账号", Width = 20, IsIgnore = false, Sort = 20, Align = TextAlign.Right, FontColor = DefineColor.LightBlue)]
		public string Account { get; set; }

		//[ExcelKit(Desc = "昵称", Width = 50, Sort = 10, FontColor = DefineColor.Rose, ForegroundColor = DefineColor.LemonChiffon)]
		public string Name { get; set; }

		//[ExcelKit(Desc = "金额", Width = 20, Sort = 10, Converter = typeof(DecimalPointDigitConverter), ConverterParam = 2)]
		public decimal Money { get; set; } = 20;

		//[ExcelKit(Desc = "创建时间", Width = 50, Sort = 10, Converter = typeof(DateTimeFmtConverter), ConverterParam = "yyyy-MM-dd")]
		public DateTime? CreateDate { get; set; }

		//[ExcelKit(Desc = "性别", Width = 50, Sort = 10, Converter = typeof(BoolConverter), ConverterParam = "√|×")]
		public bool? Sex { get; set; }

		[ExcelKit(Desc = "用户类型", Width = 20, Align = TextAlign.Center, Converter = typeof(EnumConverter<UserType>))]
		public UserType Type { get; set; } = UserType.系统用户;

		[ExcelKit(Desc = "游戏角色", Width = 20, Align = TextAlign.Center, Converter = typeof(EnumerableConverter<string>))]
		public List<string> GameRoles => new List<string>() { "射手", "法师" };
	}

	[Description("用户类型")]
	public enum UserType
	{
		[Description("系统用户")]
		系统用户 = 10,
		[Description("预置用户")]
		预置用户 = 20
	}
}
