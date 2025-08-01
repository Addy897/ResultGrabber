

from sqlalchemy import (
    Column,
    String,
    Integer,
    Float,
    CHAR,
    ForeignKey,
    CheckConstraint,
    PrimaryKeyConstraint,
    Index,
)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship

Base = declarative_base()


class Student(Base):
    __tablename__ = "Student"

    USN = Column(String(20), primary_key=True)
    StudentName = Column(String(50), nullable=False)


    results = relationship("Result", back_populates="student", cascade="all, delete-orphan")
    grades = relationship("StudentGrades", back_populates="student", uselist=False)


class Subject(Base):
    __tablename__ = "Subject"

    SubjectCode = Column(String(10), primary_key=True)
    SubjectName = Column(String(50), nullable=False)
    Semester = Column(Integer, nullable=False)


    results = relationship("Result", back_populates="subject", cascade="all, delete-orphan")
    stats = relationship("SubjectStatistics", back_populates="subject", uselist=False)


class Result(Base):
    __tablename__ = "Result"

    USN = Column(String(20), ForeignKey("Student.USN"), nullable=False)
    SubjectCode = Column(String(10), ForeignKey("Subject.SubjectCode"), nullable=False)
    SEE = Column(Integer)
    CIE = Column(Integer)
    Total = Column(Integer)
    Result = Column(CHAR(1))

    __table_args__ = (
        PrimaryKeyConstraint("USN", "SubjectCode"),
        CheckConstraint("SEE BETWEEN 0 AND 100"),
        CheckConstraint("CIE BETWEEN 0 AND 100"),
        CheckConstraint("Result IN ('P', 'F','A','NC','W','X','NE')"),
        Index("idx_usn", "USN"),
        Index("idx_subject", "SubjectCode"),
    )


    student = relationship("Student", back_populates="results")
    subject = relationship("Subject", back_populates="results")


class SubjectStatistics(Base):
    __tablename__ = "SubjectStatistics"

    SubjectCode = Column(String(10), ForeignKey("Subject.SubjectCode"), primary_key=True)
    TotalStudents = Column(Integer, default=0)
    Appeared = Column(Integer, default=0)
    Absent = Column(Integer, default=0)
    Pass = Column(Integer, default=0)
    Fail = Column(Integer, default=0)
    PassPercentage = Column(Float, default=0.0)
    FCDCount = Column(Integer, default=0)
    FCCount = Column(Integer, default=0)
    SCCount = Column(Integer, default=0)


    subject = relationship("Subject", back_populates="stats")


class StudentGrades(Base):
    __tablename__ = "StudentGrades"

    USN = Column(String(20), ForeignKey("Student.USN"), primary_key=True)
    TotalMarks = Column(Integer, default=0)
    Percentage = Column(Float, default=0.0)
    Grade = Column(String(5), CheckConstraint("Grade IN ('FCD', 'FC', 'SC')"))


    student = relationship("Student", back_populates="grades")


class SemesterStatistics(Base):
    __tablename__ = "SemesterStatistics"

    Semester = Column(Integer, primary_key=True)
    BatchYear = Column(String(10), primary_key=True)
    TotalStudents = Column(Integer, default=0)
    Pass = Column(Integer, default=0)
    Fail = Column(Integer, default=0)
    AvgTotalMarks = Column(Float, default=0.0)
    FCDCount = Column(Integer, default=0)
    FCCount = Column(Integer, default=0)
    SCCount = Column(Integer, default=0)




class OverallStatistics(Base):
    __tablename__ = "OverallStatistics"

    BatchYear = Column(String(10), primary_key=True)
    TotalStudents = Column(Integer, default=0)
    Pass = Column(Integer, default=0)
    Fail = Column(Integer, default=0)
    AvgTotalMarks = Column(Float, default=0.0)
    FCDCount = Column(Integer, default=0)
    FCCount = Column(Integer, default=0)
    SCCount = Column(Integer, default=0)





Index("idx_semester", Subject.Semester)
