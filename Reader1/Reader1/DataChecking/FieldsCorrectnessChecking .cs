﻿using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static Org.BouncyCastle.Math.EC.ECCurve;

namespace Reader1.DataChecking
{
    public class FieldsCorrectnessChecking
    {

        public static bool CheckOrganizationName(string classNumber)
        {
            return true;
        }

        public static bool CheckClassNumber(string classNumber)
        {
            if (string.IsNullOrEmpty(classNumber))
            {
                return false;
            }

            int number;
            if (int.TryParse(classNumber[0].ToString(), out number))
            {
                return (number >= 1 && number <= 11);
            }
            return false;
        }

        public static bool CheckClassLetter(string classLetter)
        {
            if (classLetter.Length != 1)
            {
                return false;
            }

            return char.IsLetter(classLetter[0]) && ((classLetter[0] >= 'а' && classLetter[0] <= 'я') || (classLetter[0] >= 'А' && classLetter[0] <= 'Я'));
        }                                                                

        public static bool CheckName(string name)
        {
            string pattern = @"^[А-Яа-яЁё]+\s[А-Яа-яЁё]+\s[А-Яа-яЁё]+$";

            Regex regex = new Regex(pattern);

            return regex.IsMatch(name);
        }

        public static bool CheckEmail(string email)
        {
            string pattern = @"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$";

            Regex regex = new Regex(pattern);

            return regex.IsMatch(email);
        }

        public static bool CheckPhone(string phone)
        {
            string pattern = @"^\+?7\d{10}$";

            Regex regex = new Regex(pattern);

            return regex.IsMatch(phone);
        }

        public static bool QuarterNumber(string num)
        {
            if (num.Length != 1)
            {
                return false;
            }
            int number;
            if (int.TryParse(num[0].ToString(), out number))
            {
                return (number >= 1 && number <= 4);
            }
            return false;
        }

        public static bool CheckYear(string num)
        {
            string pattern = @"^\d{4}$";

            Regex regex = new Regex(pattern);

            return regex.IsMatch(num);
        }
    }
}
