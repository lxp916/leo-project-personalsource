using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XPSCryptEncrypt.Lib
{
    namespace CommandExample
    {
        #region Command
        //class Program
        //{
        //    static void Main(string[] args)
        //    {
        //        Man man = new Man();//receiver
        //        Server server = new Server();//invoke
        //        server.Execute(new MoveForward(man, 10));
        //        System.Threading.Thread.Sleep(50);
        //        server.Execute(new MoveRight(man, 10));
        //        server.Execute(new MoveBackward(man, 10));
        //        server.Execute(new MoveLeft(man, 10));
        //    }
        //}


        //class Man
        //{
        //    private int x = 0;
        //    private int y = 0;


        //    public void MoveLeft(int i) { x -= i; }


        //    public void MoveRight(int i) { x += i; }


        //    public void MoveForward(int i) { y += i; }


        //    public void MoveBackward(int i) { y -= i; }


        //    public void GetLocation()
        //    {
        //        Console.WriteLine(string.Format("({0},{1})", x, y));
        //    }
        //}


        //abstract class GameCommand
        //{
        //    private DateTime time;


        //    public DateTime Time
        //    {
        //        get { return time; }
        //        set { time = value; }
        //    }


        //    protected Man man;


        //    public Man Man
        //    {
        //        get { return man; }
        //        set { man = value; }
        //    }


        //    public GameCommand(Man man)
        //    {
        //        this.time = DateTime.Now;
        //        this.man = man;
        //    }



        //    public abstract void Execute();


        //    public abstract void UnExecute();
        //}



        //class MoveLeft : GameCommand
        //{
        //    int step;


        //    public MoveLeft(Man man, int i) : base(man) { this.step = i; }


        //    public override void Execute()
        //    {
        //        man.MoveLeft(step);
        //    }


        //    public override void UnExecute()
        //    {
        //        man.MoveRight(step);
        //    }
        //}


        //class MoveRight : GameCommand
        //{
        //    int step;


        //    public MoveRight(Man man, int i) : base(man) { this.step = i; }


        //    public override void Execute()
        //    {
        //        man.MoveRight(step);
        //    }


        //    public override void UnExecute()
        //    {
        //        man.MoveLeft(step);
        //    }
        //}


        //class MoveForward : GameCommand
        //{
        //    int step;


        //    public MoveForward(Man man, int i) : base(man) { this.step = i; }


        //    public override void Execute()
        //    {
        //        man.MoveForward(step);
        //    }


        //    public override void UnExecute()
        //    {
        //        man.MoveBackward(step);
        //    }
        //}


        //class MoveBackward : GameCommand
        //{
        //    int step;


        //    public MoveBackward(Man man, int i) : base(man) { this.step = i; }


        //    public override void Execute()
        //    {
        //        man.MoveBackward(step);
        //    }


        //    public override void UnExecute()
        //    {
        //        man.MoveForward(step);
        //    }
        //}


        //class Server
        //{
        //    GameCommand lastCommand;


        //    public void Execute(GameCommand cmd)
        //    {
        //        Console.WriteLine(cmd.GetType().Name);
        //        if (lastCommand != null && (TimeSpan)(cmd.Time - lastCommand.Time) < new TimeSpan(0,
        //    0, 0, 0, 20))
        //        {
        //            Console.WriteLine("Invalid command");
        //            lastCommand.UnExecute();
        //            lastCommand = null;
        //        }
        //        else
        //        {
        //            cmd.Execute();
        //            lastCommand = cmd;
        //        }
        //        cmd.Man.GetLocation();
        //    }
        //}
        #endregion
    }

    namespace Builder_DesignPattern
    {

        //using System;

        //// These two classes could be part of a framework,

        //// which we will call DP

        //// ===============================================

        //class Director
        //{

        //    public void Construct(AbstractBuilder abstractBuilder)
        //    {

        //        abstractBuilder.BuildPartA();

        //        if (1 == 1) //represents some local decision inside director
        //        {

        //            abstractBuilder.BuildPartB();

        //        }

        //        abstractBuilder.BuildPartC();

        //    }

        //}

        //abstract class AbstractBuilder
        //{

        //    abstract public void BuildPartA();

        //    abstract public void BuildPartB();

        //    abstract public void BuildPartC();

        //}

        //// These two classes could be part of an application

        //// =================================================

        //class ConcreteBuilder : AbstractBuilder
        //{

        //    override public void BuildPartA()
        //    {

        //        // Create some object here known to ConcreteBuilder

        //        Console.WriteLine("ConcreteBuilder.BuildPartA called");

        //    }

        //    override public void BuildPartB()
        //    {

        //        // Create some object here known to ConcreteBuilder

        //        Console.WriteLine("ConcreteBuilder.BuildPartB called");

        //    }

        //    override public void BuildPartC()
        //    {

        //        // Create some object here known to ConcreteBuilder

        //        Console.WriteLine("ConcreteBuilder.BuildPartC called");
        //    }

        //}

        /////

        ///// Summary description for Client.

        /////

        //public class Client
        //{

        //    public static int Main(string[] args)
        //    {

        //        ConcreteBuilder concreteBuilder = new ConcreteBuilder();

        //        Director director = new Director();

        //        director.Construct(concreteBuilder);

        //        return 0;

        //    }

        //}

    }
}