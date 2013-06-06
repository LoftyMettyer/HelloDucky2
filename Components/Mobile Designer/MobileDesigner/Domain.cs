using System.Collections.Generic;

namespace MobileDesigner
{
    public class Entity
    {
        public virtual int Id { get; set; }
    }

    public class Layout : Entity
    {
        public virtual int HeaderBackColor { get; set; }
        public virtual Picture HeaderPicture { get; set; }
        public virtual ImageLayout HeaderPictureLocation { get; set; }
        public virtual Picture HeaderLogo { get; set; }
        public virtual int HeaderLogoWidth { get; set; }
        public virtual int HeaderLogoHeight { get; set; }
        public virtual int HeaderLogoHorizontalOffset { get; set; }
        public virtual int HeaderLogoVerticalOffset { get; set; }
        public virtual byte HeaderLogoHorizontalOffsetBehaviour { get; set; }
        public virtual byte HeaderLogoVerticalOffsetBehaviour { get; set; }
        public virtual int MainBackColor { get; set; }
        public virtual Picture MainPicture { get; set; }
        public virtual ImageLayout MainPictureLocation { get; set; }
        public virtual int FooterBackColor { get; set; }
        public virtual Picture FooterPicture { get; set; }
        public virtual ImageLayout FooterPictureLocation { get; set; }
        public virtual FontSetting TodoTitleFont { get; set; }
        public virtual int? TodoTitleForeColor { get; set; }
        public virtual FontSetting TodoDescFont { get; set; }
        public virtual int? TodoDescForeColor { get; set; }
        public virtual FontSetting HomeItemFont { get; set; }
        public virtual int? HomeItemForeColor { get; set; }
    }

    public class Element : Entity
    {
        public virtual MobileForm Form { get; set; }
        public virtual ElementType Type { get; set; }
        public virtual string Name { get; set; }
        public virtual string Caption { get; set; }
        public virtual FontSetting Font { get; set; }
        public virtual int? ForeColor { get; set; }
        public virtual Picture Picture { get; set; }
    }

    public class Picture : Entity
    {
        public virtual string Name { get; set; }
        public virtual byte[] Image { get; set; }
        public virtual PictureType Type { get; set; }

        public virtual bool Changed { get; set; }
        public virtual bool New { get; set; }
        public virtual bool Deleted { get; set; }
    }

    public class FontSetting
    {
        public string Name { get; set; }
        public float Size { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }
    }

    public class Workflow : Entity
    {
        public virtual string Name { get; set; }
        public virtual Picture Picture { get; set; }
        public virtual int InitiationType { get; set; }
        public virtual bool Enabled { get; set; }
        public virtual bool Deleted { get; set; }
    }

    public class UserGroup : Entity 
    {
        public UserGroup()
        {
            MobileWorkflows = new List<Workflow>();
        }
        public virtual string Name { get; set; }
        public virtual IList<Workflow> MobileWorkflows { get; protected set; }
    }

    public enum ElementType
    {
        Button = 0, Label = 2, TextBox = 3
    }

    public enum MobileForm
    {
        Login = 1, HomePage = 2, NewRegistration = 3, ChangePassword = 4, TodoList = 5, ForgotPassword = 6
    }

    public enum ImageLayout
    {
        TopLeft = 0, TopRight = 1, Centre = 2, LeftTile = 3, RightTile = 4, TopTile = 5, BottomTile = 6, Tile = 7
    }

    public enum ImageLocation
    {
        TopLeft = 0, TopRight = 1, BottomLeft = 2, BottomRight = 3
    }

    public enum PictureType
    {
        Bitmap = 1, Icon = 3, Png = 5
    }
}
