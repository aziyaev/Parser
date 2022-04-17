using System.Windows.Data;
using System.Collections;

namespace Parser
{
    public class PagingCollectionView : CollectionView
    {
        private IList innerList;

        private int itemsPerPage = 15;
        private int currentPage = 1;

        private string currentNote = "Идентификатор\n\nНаименование: \n\nОписание: \n\n" +
                "Источник: \n\nОбъект воздействия: \n\n" +
                "Нарушение конфиденциальности: \n\nНарушение целостности: \n\nНарушение доступности: \n\n" +
                "Дата создания: \n\nДата изменения: \n\n";

        public PagingCollectionView(IList innerList) : base(innerList)
        {
            this.innerList = innerList;
        }

        public override int Count { get
            {
                if(this.currentPage < this.PageCount)
                {
                    return this.ItemsPerPage;
                }
                else
                {
                    var itemsLeft = this.innerList.Count % this.itemsPerPage;
                    if (itemsLeft == 0)
                    {
                        return this.itemsPerPage;
                    }
                    else return itemsLeft;
                }
            } }

        public int CurrentPage { get { return this.currentPage; } 
            set 
            {
                this.currentPage = value;
                this.OnPropertyChanged(new System.ComponentModel.PropertyChangedEventArgs("CurrentPage"));
            } }

        public int ItemsPerPage { get { return this.itemsPerPage; } }

        public int PageCount
        {
            get
            {
                return (this.innerList.Count + this.itemsPerPage - 1) / this.itemsPerPage;
            }
        }

        private int EndIndex
        {
            get
            {
                var end = this.currentPage * this.itemsPerPage - 1;
                return (end > this.innerList.Count) ? this.innerList.Count : end;
            }
        }

        private int StartIndex
        {
            get { return (this.currentPage - 1) * this.itemsPerPage; }
        }
        public string CurrentNote
        {
            get { return this.currentNote; }
            set
            {
                this.currentNote = value;
                this.OnPropertyChanged(new System.ComponentModel.PropertyChangedEventArgs("CurrentNote"));
            }
        }

        public override object GetItemAt(int index)
        {
            var offset = index % (this.itemsPerPage);
            return this.innerList[this.StartIndex + offset];
        }

        public void MoveToNextPage()
        {
            if(this.currentPage < this.PageCount)
            {
                this.CurrentPage += 1;
            }
            this.Refresh();
        }

        public void MoveToPreviousPage()
        {
            if(this.currentPage > 1)
            {
                this.CurrentPage -= 1;
            }
            this.Refresh();
        }
    }
}