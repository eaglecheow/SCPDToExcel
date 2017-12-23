using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCPDRecorder.Services
{
    class InputValidator
    {
        /// <summary>
        /// Checks if the input is in the format of a Library Number
        /// </summary>
        /// <param name="input">Input got from the user</param>
        /// <returns>Is the input a valid library number</returns>
        public bool LibraryNumberValidator(string input)
        {
            bool isValid = false;
            if (input.Length == 10)
            {
                if (input.All(char.IsDigit))
                {
                    isValid = true;
                }
            }
            return isValid;
        }
    }
}
