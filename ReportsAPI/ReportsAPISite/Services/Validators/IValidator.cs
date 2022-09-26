namespace ReportsAPISite.Services.Validators
{
    public interface IValidator<T>
    {
        T Candidate { get; }
        void AddError(string message, string property);
        void ThrowIfInvalid();
    }
}