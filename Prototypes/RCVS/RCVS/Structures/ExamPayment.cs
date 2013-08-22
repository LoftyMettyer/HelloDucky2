using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using RCVS.Classes;
using RCVS.Enums;

namespace RCVS.Structures
{
	public struct ExamPayment
	{
		[Required]
		[DisplayName("The amount you are paying")]

		public double Amount { get; set; }

		[Required]
		[DisplayName("Payment method")]
		public PaymentMethod PaymentMethod { get; set; }
		public string OtherPaymentMethod { get; set; }

		[Required]
		[DisplayName("Name and address of the person who is paying the exam fee (if not the applicant)")]
		public string PayeeName { get; set; }
		public Address PayeeAddress { get; set; }

		[Required]
		[DisplayName("Payee email address")]
		public string PayeeEmail { get; set; }

		[DisplayName("I enclose the examination fee of £")]
		public double TotalAmount
		{
			get { return Amount; }
		}

	}
}