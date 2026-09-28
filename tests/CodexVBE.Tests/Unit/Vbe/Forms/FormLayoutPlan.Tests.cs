namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass,TestCategory("Unit")]
    public sealed class FormLayoutPlanTests
    {
        private static FormLayoutBox Box(string path,double left=10,double top=10,double width=10,double height=10)
        {return new FormLayoutBox{Path=path,Left=left,Top=top,Width=width,Height=height};}
        [TestMethod]
        public void LayoutValidatesCountsContainerAndEverySourceGeometryCoordinate()
        {
            foreach(var input in new[]{(FormLayoutBox[])null,new FormLayoutBox[1],new FormLayoutBox[65]})
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(input,"align_left",0,100,100));
            foreach(double invalid in new[]{double.NaN,double.PositiveInfinity,0,-1})
            {
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"align_left",0,invalid,100));
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"align_left",0,100,invalid));
            }
            foreach(var invalid in new[]{(FormLayoutBox)null,Box("a",double.NaN),Box("a",top:double.PositiveInfinity),Box("a",width:0),Box("a",height:0)})
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{invalid,Box("b")},"align_left",0,100,100));
        }
        [TestMethod]
        public void EveryAlignmentSizingCenteringAndGridActionPreservesIndependentSourceBoxes()
        {
            foreach(string action in new[]{"align_left","align_right","align_centers","align_middles","align_top","align_bottom","same_width","same_height","same_size","center_horizontal","center_vertical","snap_grid"})
            {
                var input=new[]{Box("anchor",20,20,20,20),Box("other",40,40)};
                var output=FormLayoutPlan.Create(input,action,5,100,100);
                Assert.AreEqual(40,input[1].Left); Assert.AreEqual(40,input[1].Top); Assert.AreEqual(10,input[1].Width); Assert.AreEqual(10,input[1].Height);
                Assert.AreNotSame(input[0],output[0]); Assert.AreEqual("other",output[1].Path);
                if(action=="align_left") Assert.AreEqual(20,output[1].Left);
                if(action=="align_right") Assert.AreEqual(30,output[1].Left);
                if(action=="align_centers") Assert.AreEqual(25,output[1].Left);
                if(action=="align_middles") Assert.AreEqual(25,output[1].Top);
                if(action=="align_top") Assert.AreEqual(20,output[1].Top);
                if(action=="align_bottom") Assert.AreEqual(30,output[1].Top);
                if(action=="same_width"||action=="same_size") Assert.AreEqual(20,output[1].Width);
                if(action=="same_height"||action=="same_size") Assert.AreEqual(20,output[1].Height);
                if(action=="center_horizontal") Assert.AreEqual(55,output[1].Left);
                if(action=="center_vertical") Assert.AreEqual(55,output[1].Top);
                if(action=="snap_grid") Assert.AreEqual(40,output[1].Top);
            }
            Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"unknown",0,100,100));
        }
        [TestMethod]
        public void SpacingAndDistributionCoverBothAxesAndRejectOverlapsAndInvalidSteps()
        {
            foreach(string axis in new[]{"horizontal","vertical"})
            {
                var input=new[]{Box("a",0,0),Box("b",20,20),Box("c",50,50)};
                foreach(string prefix in new[]{"space_","increase_","decrease_"})
                {
                    string action=prefix+axis+(prefix=="space_"?"":"_spacing");
                    var output=FormLayoutPlan.Create(input,action,5,100,100);
                    Assert.AreEqual(prefix=="increase_"?25:15,axis=="horizontal"?output[1].Left:output[1].Top);
                }
                var distributed=FormLayoutPlan.Create(input,"distribute_"+axis,0,100,100);
                Assert.AreEqual(25,axis=="horizontal"?distributed[1].Left:distributed[1].Top);
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"distribute_"+axis,0,100,100));
                Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",width:30,height:30),Box("b",12,12,30,30),Box("c",20,20,30,30)},"distribute_"+axis,0,100,100));
            }
            foreach(double invalid in new[]{double.NaN,double.PositiveInfinity,-1,1001})
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"space_horizontal",invalid,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",0),Box("b",11)},"decrease_horizontal_spacing",2,100,100));
        }
        [TestMethod]
        public void GridAndResultBoundsRejectNonfiniteOverflowAndOutOfContainerCoordinates()
        {
            foreach(double invalid in new[]{0,101,double.NaN})
                Assert.ThrowsException<ArgumentException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"snap_grid",invalid,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",0,width:10),Box("b",width:20)},"align_right",0,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",top:0,height:10),Box("b",height:20)},"align_bottom",0,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",95),Box("b",95)},"align_left",0,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",top:95),Box("b",top:95)},"align_top",0,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a"),Box("b")},"snap_grid",double.Epsilon,100,100));
            Assert.ThrowsException<InvalidOperationException>(()=>FormLayoutPlan.Create(new[]{Box("a",0,width:double.MaxValue),Box("b",double.MaxValue,width:double.MaxValue),Box("c",double.MaxValue,width:double.MaxValue)},"distribute_horizontal",0,double.MaxValue,100));
        }
    }
}
