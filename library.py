import pandas as pd
import xlrd
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_bootstrap import Bootstrap5
from flask_wtf import FlaskForm
from pandas.core.methods.to_dict import to_dict
from wtforms import FloatField, DecimalField, StringField, SubmitField
from wtforms.fields import StringField, SubmitField, DateTimeField, DecimalRangeField
from wtforms.validators import DataRequired, NumberRange
from flask_sqlalchemy import SQLAlchemy


db = SQLAlchemy()
app = Flask(__name__)
app.config['SECRET_KEY'] = "8BYkEfBA6O6donzW"
app.config['SQLALCHEMY_DATABASE_URI'] = "sqlite:///new_books_collection.db"
db.init_app(app)
Bootstrap5(app)


class Book(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(250), unique=True, nullable=False)
    author = db.Column(db.String(250), unique=False, nullable=False)
    series = db.Column(db.String(50), unique=False, nullable=True)
    rating = db.Column(db.Float, nullable=False)


class BookForm(FlaskForm):
    title = StringField(label='Book Name', validators=[DataRequired()], render_kw={'placeholder': 'Game of Thrones'})
    author = StringField(label='Author', validators=[DataRequired()], render_kw={'placeholder': 'George Martin'})
    series = StringField(label='Series', render_kw={'placeholder': 'The trial of fire series'})
    rating = FloatField(label="Rating",
                        validators=[DataRequired(), NumberRange(min=0.0, max=10.0, message="Enter Between 0.0-10.0")],
                        render_kw={'placeholder': '6.8'})
    submit = SubmitField(label='Add Book')


class UpdateForm(FlaskForm):
    title = StringField(label='Book Name', validators=[DataRequired()])
    author = StringField(label='Author', validators=[DataRequired()])
    series = StringField(label='Series')
    rating = FloatField(label="Rating",
                        validators=[DataRequired(), NumberRange(min=0.0, max=10.0, message="Enter Between 0.0-10.0")],
                        render_kw={'placeholder': '6.8'})
    submit = SubmitField(label='Update Book Details')


class SearchForm(FlaskForm):
    title = StringField(label='Book Name')
    author = StringField(label='Author')
    submit = SubmitField(label='Search for Books')


def to_dict(row):
    if row is None:
        return None

    rtn_dict = dict()
    keys = row.__table__.columns.keys()
    for key in keys:
        rtn_dict[key] = getattr(row, key)
    return rtn_dict


##Create Table
with app.app_context():
    db.create_all()


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/add", methods=["GET", "POST"])
def add():
    form = BookForm()
    result = db.session.execute(db.select(Book).order_by(Book.title))
    all_db_books = result.scalars()
    if form.validate_on_submit():
        new_book = Book(
            title=request.form['title'],
            author=request.form['author'],
            series=request.form['series'],
            rating=request.form['rating']
        )
        for item in all_db_books:
            if item.title == new_book.title:
                flash(message=f"Book Name : {new_book.title} already exists. Please check information entered.",
                      category="alert-success")
                return redirect(url_for('add'))
        db.session.add(new_book)
        db.session.commit()
        flash(message=f"Book Name : {new_book.title} Added successfully", category="alert-success")
        return redirect(url_for('home'))
    return render_template('add.html', form=form)


@app.route("/book/<int:book_id>", methods=["GET", "POST"])
def update(book_id):
    form1 = UpdateForm()
    book_to_edit = db.get_or_404(Book, book_id)
    if request.method == "GET":
        form1.title.data = book_to_edit.title
        form1.author.data = book_to_edit.author
        form1.rating.data = book_to_edit.rating
    if form1.validate_on_submit():
        book_to_edit.author = form1.author.data
        book_to_edit.title = form1.title.data
        book_to_edit.rating = form1.rating.data
        db.session.commit()
        flash(f"Rating for book: {book_to_edit.title} updated successfully", category="alert-success")
        return redirect(url_for('home'))
    # print(book_to_edit.title)
    # flash(book_to_edit)
    return render_template('update.html', form=form1)


@app.route("/delete")
def delete():
    book_id = request.args.get("book_id")
    book_to_delete = db.get_or_404(Book, book_id)
    db.session.delete(book_to_delete)
    db.session.commit()
    flash(message=f"Book Name : {book_to_delete.title} Deleted successfully", category="alert-danger")
    return redirect(url_for('home'))


@app.route("/list", methods=["GET", "POST"])
def listing():
    result = db.session.execute(db.select(Book).order_by(Book.title))
    all_db_books = result.scalars()
    return render_template("list.html", books=all_db_books)



@app.route("/search", methods=["GET", "POST"])
def search():
    form2 = SearchForm()
    if form2.validate_on_submit():
        title = form2.title.data
        author = form2.author.data
        result = db.session.execute(db.select(Book).order_by(Book.title))
        all_db_books = result.scalars()
        return render_template("searched.html", books=all_db_books, title=title, author=author )
    return render_template("search.html", form=form2)


@app.route("/export", methods=['GET', 'POST'])
def export():

    result = db.session.execute(db.select(Book).order_by(Book.title))
    all_db_books = result.scalars()

    cols = ["id", "title", "author", "series", "rating"]
    data = Book.query.all()
    data_list = [to_dict(item) for item in data]
    df = pd.DataFrame(data_list)

    filename = "static/data.xlsx"
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, sheet_name="Books")
    writer._save()
    flash(f"Export Successful! Your database has been save under {filename}", category="alert-success")
    return redirect(url_for('home'))


@app.route("/delete_db")
def delete_db():
    # flash(f"delete database", category="alert-success")
    # return redirect(url_for('home'))
    return render_template("delete_db.html")


@app.route("/delete_db1")
def delete_db1():
    db.session.query(Book).delete()
    db.session.commit()
    flash(f"Database Deletion Successful!")
    return redirect(url_for('home'))


@app.route("/import")
def import_db():
    result = db.session.execute(db.select(Book).order_by(Book.title))
    all_db_books = result.scalars()

    # https://pandas.pydata.org/pandas-docs/stable/user_guide/io.html#io-ods
    xls = xlrd.open_workbook("static/test.xls", on_demand=True)
    worksheet = xls.sheet_by_index(0)
    first_row = []
    for col in range(worksheet.ncols):
        first_row.append(worksheet.cell_value(0, col))
    # transform the workbook to a list of dictionaries
    data = []

    for row in range(1, worksheet.nrows):
        elm = {}
        for col in range(worksheet.ncols):
            elm[first_row[col]] = worksheet.cell_value(row, col)
        data.append(elm)
    # print(data)
    for item in data:
        author1 = item["author"]
        title1 = item["title"]
        series1 = item["series"]
        rating1 = float(item["rating"])
        new_book = Book(
            title=title1,
            author=author1,
            series=series1,
            rating=rating1
                        )
        for book in all_db_books:
            if book.title == new_book.title:
                flash(message=f"Book Name : {new_book.title} already exists. Please remove from database "
                              f"and try again.", category="alert-success")
                return redirect(url_for('home'))
        db.session.add(new_book)
        db.session.commit()

    flash(f"Import Database Successfully", category="alert-success")
    return redirect(url_for('home'))


if __name__ == "__main__":
    app.run(debug=True)
