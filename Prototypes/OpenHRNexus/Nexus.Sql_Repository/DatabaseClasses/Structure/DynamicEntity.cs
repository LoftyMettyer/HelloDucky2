using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

public abstract class DynamicEntity : INotifyPropertyChanged
{
    public abstract int Id { get; set; }
    public abstract event PropertyChangedEventHandler PropertyChanged;
    public abstract void OnPropertyChanged(string propertyName);
}

public class BaseDynamicEntity : DynamicEntity
{
    private int _id;

    [DataMember, Key]
    public override int Id
    {
        get
        {
            return _id;
        }
        set
        {
            _id = value;
            OnPropertyChanged("Id");
        }
    }

    public override event PropertyChangedEventHandler PropertyChanged;

    public override void OnPropertyChanged(string propertyName)
    {
        if (PropertyChanged != null)
        {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}